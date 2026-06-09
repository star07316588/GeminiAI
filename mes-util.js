// 廣域常數
window.commonError = "Fail to execute application, please call IT support!!\n程式執行失敗, 請洽IT人員處理";

// 共用函數 checkAuth，供頁面呼叫驗證登入狀態
window.checkAuth = async function () {
    let relogin = false;
    const token = sessionStorage.getItem('mes_token')
    if (!token) {
      relogin = true;
    } else if (isSessionExpired()) {
        relogin = true;
        console.log('Session 已過期，請重新登入');
      if (typeof VMessageBox !== 'undefined' && VMessageBox.alert && typeof VMessageBox.alert === 'function') {
        VMessageBox.alert('Session 已過期，請重新登入');
      }
      else {
        alert ('Session 已過期，請重新登入');
      }
    }
    // 可執行登出流程或跳轉登入頁
    if (relogin) {
        sessionStorage.setItem('mes_lasturl', window.location.href);
        //window.location.replace('/Content/login.html');
        window.location.replace((window.root || '/') + 'Home/Login');
        return false;
    }

    return {
        token,
        userName: sessionStorage.getItem('mes_loginId') || '未登入'
    }
};

// 共用函數 mesfetch
/**
 * 共用的自訂 Fetch 函數 mesfetch，預設為 POST 與 Content-Type: application/json
 * @param {string} url - 請求的 URL
 * @param {object} options - fetch 的設定選項，可覆寫 method、headers、body 等
 * @returns {Promise<object>} - 回傳解析後的 JSON 資料
 */
async function mesfetch(url, body, options = {}) {
  // 預設選項
  const defaultOptions = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${sessionStorage.getItem('mes_token')}`
    },
    body: JSON.stringify(body)
  };

  // 合併 headers，確保有 Content-Type 且可以覆寫
  const mergedHeaders = {
    ...defaultOptions.headers,
    ...(options.headers || {})
  };

  // 最終 options，options 會覆寫 defaultOptions，但 headers 合併
  const finalOptions = {
    ...defaultOptions,
    ...options,
    headers: mergedHeaders
  };

  // 若 method 為 GET 移除 body
  if (finalOptions.method.toUpperCase() === 'GET') {
    delete finalOptions.body;
  }

  try {
    const response = await fetch(url, finalOptions);

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const data = await response.json();
    return data;
  } catch (error) {
    console.error('mesfetch error:', error);
    throw error;
  }
}

// 檢查登入是否過期 (現在是半小時)
function isSessionExpired() {
  const loginTimeStr = sessionStorage.getItem('loginTime');
  if (!loginTimeStr) {
    console.log('Session 沒有 loginTime，視為過期');
    return true;
  }
  console.log('loginTime:', loginTimeStr);
  const loginTime = parseInt(loginTimeStr, 10);
  const now = Date.now();
  const thirtyMinutes = 30 * 60 * 1000; // 30 分鐘的毫秒數

  return now - loginTime > thirtyMinutes;
}


// mes-toolbar
Vue.component('mes-toolbar', {
  props: {
    title: { type: String, default: '' },
    //user:  { type: String, default: '未登入' },
    items: { type: Array,  default: () => [] }
  },
  data() {
    return {
      dropdownOpen: false,
      user: sessionStorage.getItem('mes_loginId') || '未登入'
    }
  },
  methods: {
    toggleDropdown() {
      this.dropdownOpen = !this.dropdownOpen;
    },
    onChangePassword() {
      this.$emit('change-password');
      this.dropdownOpen = false;
    },
    onLogout() {
      this.$emit('logout');
      this.dropdownOpen = false;
    },
    closeDropdown() {
      this.dropdownOpen = false;
    }
  },
  mounted() {
    document.addEventListener('click', this.handleClickOutside);
  },
  beforeDestroy() {
    document.removeEventListener('click', this.handleClickOutside);
  },
  methods: {
    toggleDropdown() {
      this.dropdownOpen = !this.dropdownOpen;
    },
    handleClickOutside(event) {
      if (!this.$el.contains(event.target)) {
        this.dropdownOpen = false;
      }
    },
    onChangePassword() {
      //this.$emit('change-password');
      this.dropdownOpen = false;
//      ElMessageBox.alert('系統尚未完成');
    },
    async onLogout() {
//      this.$emit('logout');
        this.dropdownOpen = false;

        let comfired = false;
        if (typeof VMessageBox !== 'undefined' && VMessageBox.alert && typeof VMessageBox.alert === 'function') {
          var res = await VMessageBox.confirm('確定要登出嗎？');
          comfired = res?.action == 'confirm';
        }
        else {
          comfired = confirm('確定要登出嗎？');
        }
        if (comfired) {
            console.log('執行登出...');
            sessionStorage.clear();
            window.location.href = '/Content/login.html';
        }        

    }
  },
  template: `
    <div class="mes-toolbar">
      <div class="tb-logo">
        <div class="tb-logo-hex"></div>
        <span class="tb-logo-text">BE MES</span>
      </div>

      <span v-if="title" class="tb-title"
            style="padding-right:14px; border-right:1px solid var(--border); margin-right:6px;">
        {{ title }}
      </span>

      <!-- 可擴充功能鈕 -->
      <div class="tb-group" v-if="items.length">
        <div v-for="(item, i) in items" :key="i"
             class="tb-btn" :title="item.label"
             @click="$emit('item-click', item.action)">
          <i v-if="item.icon" :class="'mdi ' + item.icon + ' tb-btn-icon'"></i>
          <span class="tb-btn-label">{{ item.label }}</span>
        </div>
      </div>

      <div class="tb-spacer"></div>

      <!-- 使用者欄位，點擊展開下拉選單 -->
      <div class="tb-right" style="position:relative;">
        <div class="tb-user" @click="toggleDropdown" title="點擊展開選單" style="cursor:pointer; user-select:none;">
          <i class="mdi mdi-account-circle" style="color:var(--text-dim);font-size:16px;"></i>
          <span class="tb-user-name">{{ user || '—' }}</span>
          <i class="mdi mdi-chevron-down" style="color:var(--text-muted);font-size:13px;"></i>
        </div>

        <!-- Dropdown menu -->
        <div v-if="dropdownOpen" class="tb-user-dropdown" 
             style="position:absolute; right:0; top:calc(100% + 6px); background: var(--bg-surface); border: 1px solid var(--border); border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15); min-width:140px; z-index: 1000;">
          <div class="dropdown-item"
               @click="onChangePassword"
               style="padding: 8px 12px; cursor:pointer; font-size:13px; border-bottom:1px solid var(--border);">
            <i class="mdi mdi-lock-outline" style="margin-right:6px; vertical-align:middle;"></i> 更改密碼
          </div>
          <div class="dropdown-item"
               @click="onLogout"
               style="padding: 8px 12px; cursor:pointer; font-size:13px;">
            <i class="mdi mdi-logout" style="margin-right:6px; vertical-align:middle;"></i> 登出
          </div>
        </div>
      </div>
    </div>
  `
});

// mes-statusbar
Vue.component('mes-statusbar', {
  props: {
    items:   { type: Array,  default: () => [] },
    msg:     { type: String, default: '' },
    msgType: { type: String, default: '' }
  },
  data () { return { clock: '' } },
  mounted () {
    this.tick()
    this._t = setInterval(this.tick, 1000)
  },
  beforeDestroy () { clearInterval(this._t) },
  methods: {
    tick () {
      const d = new Date(), p = n => String(n).padStart(2, '0')
      this.clock = `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`
    }
  },
  template: `
    <div class="status-bar">
      <!-- 左側：可擴充項目 -->
      <template v-for="(item, i) in items">
        <div class="sb-item" :key="'item-'+i">
          <i v-if="item.icon" :class="'mdi ' + item.icon" style="font-size:11px;"></i>
          {{ item.label }}
          <span class="sb-val" :class="item.cls || ''">{{ item.value }}</span>
        </div>
        <div class="sb-divider" :key="'sep-'+i"></div>
      </template>

      <div class="sb-spacer"></div>

      <!-- 狀態訊息 -->
      <span v-if="msg" class="sb-msg" :class="msgType">{{ msg }}</span>
      <div v-if="msg" class="sb-divider"></div>

      <!-- 固定：日期時間 -->
      <div class="sb-item">
        <i class="mdi mdi-clock-outline" style="font-size:11px;"></i>
        <span class="sb-val">{{ clock }}</span>
      </div>
    </div>
  `
});

/*!
 * v-message-box.js  —  VMessageBox Plugin for Vue 2
 * ─────────────────────────────────────────────────────────────
 * 前置依賴（須在本檔案之前載入）：
 *   - Vue 2.x       (window.Vue)
 *   - Vuetify 2.x   (window.Vuetify)  — 用於 <v-icon>
 *   - @mdi/font CSS
 *   - v-message-box.css
 *
 * ── 兩種使用方式，任選其一 ───────────────────────────────────
 *
 * 方式 A｜Vue Plugin（在 Vue 實例內用 this.$xxx）
 *   Vue.use(VMessageBox)
 *   → this.$alert / this.$confirm / this.$prompt
 *
 * 方式 B｜獨立呼叫（完全不需要 Vue 實例）
 *   VMessageBox.alert('訊息')
 *   VMessageBox.confirm('訊息', '標題')
 *   VMessageBox.prompt('訊息', '標題', { inputValidator: v => !v ? '必填' : null })
 *
 * ── Promise 回傳值 ───────────────────────────────────────────
 *   resolve → { action: 'confirm', value: string }
 *   reject  → { action: 'cancel',  value: null   }
 *
 * ── opts 可用欄位 ────────────────────────────────────────────
 *   confirmText      {string}
 *   cancelText       {string}
 *   icon             {string}   MDI icon，e.g. 'mdi-check-circle-outline'
 *   color            {string}   主題色（CSS color）
 *   iconBg           {string}   icon 背景色
 *   closeOnBackdrop  {boolean}  點背景關閉（alert 預設 false，其餘 true）
 *   inputPlaceholder {string}   prompt placeholder
 *   inputValidator   {Function} prompt 驗證，回傳錯誤字串或 null
 */
(function (global, factory) {
  typeof module !== 'undefined' && module.exports
    ? (module.exports = factory())
    : (global.VMessageBox = factory());
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  'use strict';

  /* ── Dialog template ─────────────────────────────────────── */
  const TEMPLATE = `
    <transition name="vmb">
      <div class="vmb-overlay" v-if="box.visible" @mousedown.self="onBackdropClick">
        <div class="vmb-backdrop" @mousedown.self="onBackdropClick"></div>

        <div class="vmb-card"
             :style="{ '--vmb-bar-color': box.color }"
             role="dialog"
             aria-modal="true"
             :aria-labelledby="'vmb-title-' + _uid">

          <div class="vmb-body">
            <div class="vmb-icon-wrap"
                 v-if="box.icon"
                 :style="{ background: box.iconBg, color: box.color }">
              <v-icon :color="box.color">{{ box.icon }}</v-icon>
            </div>

            <h2 class="vmb-title" v-if="box.title" :id="'vmb-title-' + _uid">{{ box.title }}</h2>
            <p  class="vmb-message" v-html="box.message"></p>

            <div class="vmb-input-wrap" v-if="box.type === 'prompt'">
              <input
                ref="promptInput"
                class="vmb-input"
                :class="{ 'vmb-input--error': box.inputError }"
                v-model="box.inputValue"
                :placeholder="box.inputPlaceholder"
                @keyup.enter="onConfirm"
                @keyup.esc="onCancel"
              />
              <p class="vmb-input-error-msg" v-if="box.inputError">{{ box.inputError }}</p>
            </div>
          </div>

          <div class="vmb-footer">
            <button v-if="box.showCancel"
              class="vmb-action vmb-btn-cancel"
              @click="onCancel">{{ box.cancelText }}</button>
            <button
              class="vmb-action vmb-btn-confirm"
              :style="{ background: box.color }"
              @click="onConfirm">{{ box.confirmText }}</button>
          </div>
        </div>
      </div>
    </transition>
  `;

  /* ── Default state factory ───────────────────────────────── */
  function defaultState() {
    return {
      visible:          false,
      type:             'alert',
      title:            null,
      message:          '',
      confirmText:      '確定',
      cancelText:       '取消',
      showCancel:       false,
      closeOnBackdrop:  true,
      icon:             'mdi-information-outline',
      color:            '#5865f2',
      iconBg:           'rgba(88,101,242,.12)',
      inputValue:       '',
      inputPlaceholder: '請輸入...',
      inputValidator:   null,
      inputError:       '',
      _resolve:         null,
      _reject:          null,
    };
  }

  /* ── Core：建立響應式狀態 + 掛載獨立 Vue 實例 ──────────────
     此函式只執行一次（單例），可由 install() 或 _ensureMount()
     在任意時機觸發。                                           */
  let _state  = null;   // Vue.observable 響應式狀態
  let _open   = null;   // open(options) → Promise

  function _bootstrap(Vue) {
    if (_state) return;   // 已初始化，跳過

    // 確保 Vuetify 已向 Vue 註冊（冪等，重複呼叫無害）
    if (typeof Vuetify !== 'undefined') Vue.use(Vuetify); // eslint-disable-line no-undef

    _state = Vue.observable(defaultState());

    _open = function (options) {
      return new Promise((resolve, reject) => {
        Object.assign(
          _state,
          defaultState(),
          options,
          { _resolve: resolve, _reject: reject, visible: true }
        );
      });
    };

    /* 獨立掛載 Dialog 根節點，與宿主 #app 無關 */
    const container = document.createElement('div');
    document.body.appendChild(container);

    const vuetifyInstance = typeof Vuetify !== 'undefined' ? new Vuetify() : undefined; // eslint-disable-line no-undef
    new Vue({
      vuetify: vuetifyInstance,
      template: TEMPLATE,
      data() { return { box: _state }; },
      watch: {
        'box.visible'(v) {
          if (v && this.box.type === 'prompt') {
            this.$nextTick(() => {
              this.$refs.promptInput && this.$refs.promptInput.focus();
            });
          }
        },
      },
      methods: {
        onBackdropClick() {
          if (this.box.closeOnBackdrop) this.onCancel();
        },
        onCancel() {
          this.box.visible = false;
          this.box._reject({ action: 'cancel', value: null });
        },
        onConfirm() {
          if (this.box.type === 'prompt') {
            const v = this.box.inputValue;
            if (this.box.inputValidator) {
              const err = this.box.inputValidator(v);
              if (err) { this.box.inputError = err; return; }
            } else if (!v || !v.trim()) {
              this.box.inputError = '此欄位不能為空'; return;
            }
          }
          const value = this.box.inputValue;
          this.box.visible = false;
          this.box._resolve({ action: 'confirm', value });
        },
      },
      mounted() {
        window.addEventListener('keydown', e => {
          if (!this.box.visible) return;
          if (e.key === 'Escape') this.onCancel();
          if (e.key === 'Enter' && this.box.type !== 'prompt') this.onConfirm();
        });
      },
    }).$mount(container);
  }

  /* 在呼叫 API 時若尚未 bootstrap，自動以 window.Vue 初始化 */
  function _ensureMount() {
    if (_open) return;
    if (typeof Vue === 'undefined') {
      throw new Error('[VMessageBox] 找不到 Vue，請確認已在本檔案之前載入 Vue 2。');
    }
    _bootstrap(Vue); // eslint-disable-line no-undef
  }

  /* ── Public API ──────────────────────────────────────────── */
  const VMessageBox = {

    /* ── 方式 A：Vue Plugin install hook ─────────────────── */
    install(Vue) {
      if (VMessageBox._installed) return;
      VMessageBox._installed = true;
      _bootstrap(Vue);
      /* 掛載至 Vue prototype，讓所有實例可用 this.$xxx */
      Vue.prototype.$msgbox   = VMessageBox;
      Vue.prototype.$alert    = VMessageBox.alert.bind(VMessageBox);
      Vue.prototype.$confirm  = VMessageBox.confirm.bind(VMessageBox);
      Vue.prototype.$prompt   = VMessageBox.prompt.bind(VMessageBox);
    },

    /* ── 方式 B：獨立呼叫（完全不需要 Vue 實例） ─────────── */
    alert(message, title = null, opts = {}) {
      _ensureMount();
      return _open({
        type:            'alert',
        title:           title,
        message,
        showCancel:      false,
        closeOnBackdrop: false,
        icon:        opts.icon        ?? null,
        color:       opts.color       ?? '#5865f2',
        iconBg:      opts.iconBg      ?? 'rgba(88,101,242,.12)',
        confirmText: opts.confirmText ?? '確定',
        ...opts,
      });
    },

    confirm(message, title = null, opts = {}) {
      _ensureMount();
      return _open({
        type:        'confirm',
        title:       title,
        message,
        showCancel:  true,
        icon:        opts.icon        ?? null,
        color:       opts.color       ?? '#5865f2',
        iconBg:      opts.iconBg      ?? 'rgba(88,101,242,.12)',
        confirmText: opts.confirmText ?? '確定',
        cancelText:  opts.cancelText  ?? '取消',
        ...opts,
      });
    },

    prompt(message, title = null, opts = {}) {
      _ensureMount();
      return _open({
        type:             'prompt',
        title:            title,
        message,
        showCancel:       true,
        icon:             opts.icon             ?? null,
        color:            opts.color            ?? '#1a1d2e',
        iconBg:           opts.iconBg           ?? 'rgba(26,29,46,.08)',
        confirmText:      opts.confirmText      ?? '確定',
        cancelText:       opts.cancelText       ?? '取消',
        inputPlaceholder: opts.inputPlaceholder ?? '請輸入...',
        inputValidator:   opts.inputValidator   ?? null,
        ...opts,
      });
    },
  };

  return VMessageBox;
});
