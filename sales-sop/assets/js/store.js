/*!
 * store.js — 销售&跟单工作平台 数据层
 * 数据保存在浏览器 localStorage，支持全量 JSON 导入/导出做备份与跨人同步。
 */
(function (root) {
  'use strict';

  var KEY = 'qs_sop_v1';

  /** 角色定义（按 SOP 的层级顺序：主管理员 > 团队主管 > 销售 > 跟单） */
  var ROLES = {
    admin: { label: '主管理员', level: 4 },
    manager: { label: '团队主管', level: 3 },
    sales: { label: '销售', level: 2 },
    followup: { label: '跟单', level: 1 },
  };

  var TASK_STATUS = { open: '未完成', done: '已完成' };

  var DEFAULT_NAV = [
    { id: 'todo', label: '日常待办事项', icon: '☑', children: [] },
    { id: 'orders', label: '大货订单跟进', icon: '▤', children: [] },
    { id: 'assign', label: '任务指派', icon: '⇄', children: [] },
    { id: 'export', label: '每日跟进记录导出', icon: '⤓', children: [] },
    {
      id: 'data',
      label: '数据源表格',
      icon: '▦',
      children: [
        { id: 'data-customers', label: '客户' },
        { id: 'data-personnel', label: '人员架构表' },
        { id: 'data-accounts', label: '账号登录' },
      ],
    },
  ];

  var state = null;

  /* ---------------- 工具 ---------------- */

  /** 轻量哈希：仅用于避免明文比对，不构成安全防护 */
  function hash(str) {
    var h = 0x811c9dc5;
    str = String(str);
    for (var i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
    }
    return ('00000000' + h.toString(16)).slice(-8);
  }

  function uid(prefix) {
    return (prefix || 'x') + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  }

  function today() {
    var d = new Date();
    return (
      d.getFullYear() +
      '-' +
      String(d.getMonth() + 1).padStart(2, '0') +
      '-' +
      String(d.getDate()).padStart(2, '0')
    );
  }

  function nowStamp() {
    var d = new Date();
    return (
      today() +
      ' ' +
      String(d.getHours()).padStart(2, '0') +
      ':' +
      String(d.getMinutes()).padStart(2, '0') +
      ':' +
      String(d.getSeconds()).padStart(2, '0')
    );
  }

  /** 把 "2026年09月28日 15:38" 之类归一化为 "2026-09-28 15:38" */
  function normalizeDate(s) {
    if (!s) return '';
    s = String(s).trim().replace(/年|月/g, '-').replace(/日/g, '');
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s*(\d{1,2}:\d{2})?/);
    if (!m) return String(s).trim();
    var out = m[1] + '-' + String(+m[2]).padStart(2, '0') + '-' + String(+m[3]).padStart(2, '0');
    return m[4] ? out + ' ' + m[4] : out;
  }

  function dateOnly(s) {
    if (!s) return '';
    var m = String(s).match(/^(\d{4}-\d{2}-\d{2})/);
    return m ? m[1] : '';
  }

  function blankState() {
    return {
      version: 1,
      meta: { createdAt: nowStamp(), initialized: false },
      teams: [],
      personnel: [],
      accounts: [],
      customers: [],
      tasks: [],
      orders: [],
      navOrder: DEFAULT_NAV,
      session: null,
      settings: { companyName: '销售&跟单工作平台', subtitle: 'QUICK SHOW 捷展 · 潮牌供应链' },
    };
  }

  /* ---------------- 读写 ---------------- */

  function save() {
    try {
      localStorage.setItem(KEY, JSON.stringify(state));
      return true;
    } catch (e) {
      console.error('[store] 保存失败', e && e.message);
      return false;
    }
  }

  function load() {
    state = null;
    try {
      var raw = localStorage.getItem(KEY);
      if (raw) state = JSON.parse(raw);
    } catch (e) {
      console.error('[store] 读取失败，将重建', e && e.message);
    }
    if (!state || typeof state !== 'object') state = blankState();
    // 补齐缺省字段（向前兼容）
    var base = blankState();
    for (var k in base) {
      if (!(k in state)) state[k] = base[k];
    }
    if (!Array.isArray(state.navOrder) || !state.navOrder.length) state.navOrder = DEFAULT_NAV;
    if (!state.settings) state.settings = base.settings;
    return state;
  }

  /** 首次运行：写入种子数据（客户 + 人员），但不预置任何账号 */
  function seedIfEmpty() {
    if (state.meta.initialized) return false;
    if (root.SEED_CUSTOMERS && !state.customers.length) {
      state.customers = root.SEED_CUSTOMERS.slice();
    }
    if (root.SEED_PERSONNEL && !state.personnel.length) {
      state.personnel = root.SEED_PERSONNEL.slice();
    }
    if (!state.teams.length) {
      var hasSales = state.personnel.some(function (p) {
        return p.team;
      });
      var names = [];
      state.personnel.forEach(function (p) {
        if (p.team && names.indexOf(p.team) === -1) names.push(p.team);
      });
      state.teams = (names.length ? names : ['销售部']).map(function (n) {
        return { id: uid('t'), name: n };
      });
    }
    return true;
  }

  /* ---------------- 会话与权限 ---------------- */

  function currentUser() {
    if (!state.session) return null;
    var login = state.session.login;
    return (
      state.accounts.find(function (a) {
        return a.login === login;
      }) || null
    );
  }

  function signIn(login, password) {
    login = String(login || '').trim().toLowerCase();
    var acc = state.accounts.find(function (a) {
      return a.login === login;
    });
    if (!acc) return { error: '账号不存在，请联系主管理员开通' };
    if (!acc.active) return { error: '该账号已停用' };
    if (acc.pwHash !== hash(password)) return { error: '密码不正确（密码为工作手机号）' };
    state.session = { login: acc.login, at: nowStamp() };
    save();
    return { data: acc };
  }

  function signOut() {
    state.session = null;
    save();
  }

  function roleOf(user) {
    return (user && user.role) || 'followup';
  }

  function isAdmin(user) {
    return roleOf(user) === 'admin';
  }

  /** 指派权限：可指派给同级及下级；跟单可额外指派给销售（SOP 明确要求） */
  function canAssign(user, target) {
    if (!user || !target) return false;
    var me = ROLES[roleOf(user)];
    var you = ROLES[roleOf(target)];
    if (!me || !you) return false;
    if (roleOf(user) === 'admin') return true;
    if (roleOf(user) === 'manager') return you.level <= me.level && target.team === user.team;
    if (roleOf(user) === 'sales') return roleOf(target) === 'sales' || roleOf(target) === 'followup';
    if (roleOf(user) === 'followup') return roleOf(target) === 'sales' || roleOf(target) === 'followup';
    return false;
  }

  /** 当前用户可指派的人员列表 */
  function assignableTargets(user) {
    return state.personnel.filter(function (p) {
      if (!p.active) return false;
      if (p.login === user.login) return true;
      return canAssign(user, p);
    });
  }

  /* ---------------- 客户 ---------------- */

  function customerKey(name) {
    return String(name || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '');
  }

  function addCustomer(rec) {
    var name = String(rec.name || '').trim();
    if (!name) return { error: '客户名称不能为空' };
    var exists = state.customers.some(function (c) {
      return customerKey(c.name) === customerKey(name);
    });
    if (exists) return { error: '客户「' + name + '」已存在，未重复新建' };
    var item = Object.assign({ id: uid('c'), createdAt: today(), origin: '手动新增' }, rec, { name: name });
    state.customers.push(item);
    save();
    return { data: item };
  }

  function updateCustomer(id, patch) {
    var c = state.customers.find(function (x) {
      return x.id === id;
    });
    if (!c) return { error: '客户不存在' };
    Object.assign(c, patch);
    save();
    return { data: c };
  }

  function removeCustomer(id) {
    var i = state.customers.findIndex(function (x) {
      return x.id === id;
    });
    if (i < 0) return { error: '客户不存在' };
    state.customers.splice(i, 1);
    save();
    return { data: true };
  }

  /**
   * 导入 CRM 客户表（按名称去重，只新增没有的）
   * @param {Array<Object>} rows 已解析的客户对象数组
   */
  function importCustomers(rows) {
    var seen = {};
    state.customers.forEach(function (c) {
      seen[customerKey(c.name)] = true;
    });
    var added = 0;
    var skipped = 0;
    var buffer = [];
    rows.forEach(function (r) {
      var name = String(r.name || '').trim();
      if (!name) return;
      var k = customerKey(name);
      if (seen[k]) {
        skipped++;
        return;
      }
      seen[k] = true;
      buffer.push(
        Object.assign({ id: uid('c'), createdAt: today(), origin: 'CRM导入' }, r, { name: name })
      );
      added++;
    });
    state.customers = state.customers.concat(buffer);
    save();
    return { data: { added: added, skipped: skipped, total: state.customers.length } };
  }

  /**
   * CRM 客户表的表头别名。导入时按「精确匹配优先、包含匹配兜底」识别列。
   * 顺序即优先级。
   */
  var CUSTOMER_HEADERS = {
    name: ['*名称', '客户名称', '名称', '客户'],
    id: ['ID', '客户ID'],
    createdAt: ['创建日期'],
    source: ['客户来源', '来源'],
    ctype: ['客户类型', '类型'],
    cmethod: ['客户对接方式', '对接方式'],
    sales: ['业务员', '销售'],
    level: ['客户级别', '级别'],
    lastFollow: ['最新跟进日期', '跟进日期'],
    lastContent: ['最新跟进内容', '跟进内容'],
    address: ['地址'],
    wechat: ['微信'],
    decision: ['决策人'],
    tags: ['客户标签', '标签'],
    phone: ['电话', '手机'],
    brand: ['品牌'],
    style: ['产品风格', '风格'],
    channel: ['销售渠道', '渠道'],
    pattern: ['业务模式'],
    inquiry: ['询盘内容'],
    trade: ['交易方式'],
    pricePref: ['价格偏好'],
    // 能力与订单数据（CRM 表里带上了就一并留存）
    isNew: ['当前年份新客', '新客'],
    production: ['生产能力'],
    design: ['设计能力'],
    salesPower: ['销售能力'],
    o7Count: ['近7天订单件数'],
    o7Amount: ['近7天订单金额'],
    o30Count: ['近30天订单件数'],
    o30Amount: ['近30天订单金额'],
    oLmCount: ['上月订单件数'],
    oLmAmount: ['上月订单金额'],
  };

  // 预计算匹配表：精确匹配用字典，模糊匹配按候选词长度降序（长词优先，避免「销售」吃掉「销售渠道」）
  var _exactMap = null;
  var _looseCands = null;
  function buildHeaderIndex() {
    if (_exactMap) return;
    _exactMap = {};
    _looseCands = [];
    Object.keys(CUSTOMER_HEADERS).forEach(function (key) {
      CUSTOMER_HEADERS[key].forEach(function (text) {
        if (!(text in _exactMap)) _exactMap[text] = key;
        _looseCands.push({ text: text, key: key });
      });
    });
    _looseCands.sort(function (a, b) {
      return b.text.length - a.text.length;
    });
  }

  function normHeader(h) {
    return String(h == null ? '' : h).trim();
  }

  function exactHeader(h) {
    buildHeaderIndex();
    return _exactMap[normHeader(h)] || null;
  }

  /** 最长候选优先的包含匹配 */
  function looseHeader(h, used) {
    buildHeaderIndex();
    h = normHeader(h);
    if (!h) return null;
    for (var i = 0; i < _looseCands.length; i++) {
      var c = _looseCands[i];
      if (used && used[c.key]) continue;
      if (h.indexOf(c.text) !== -1) return c.key;
    }
    return null;
  }

  /**
   * 把一个表头单元格映射到客户字段名，识别不了返回 null。
   * 先做精确匹配，再做（长词优先的）包含匹配。
   */
  function mapCustomerHeader(h) {
    return exactHeader(h) || looseHeader(h, null);
  }

  /**
   * 把「表头行 + 数据行」的二维数组转成客户对象数组。
   * 每个字段只会被认领一次，避免同名列相互覆盖。
   * @param {Array<Array<string>>} rows 含表头在内的全部行
   */
  function mapCustomerRows(rows) {
    if (!rows || !rows.length) return { payload: [], mapping: [], rows: 0, skippedNoName: 0 };
    var header = rows[0] || [];
    var mapping = [];
    var used = {};
    var i, k;

    // 第一轮：精确匹配
    for (i = 0; i < header.length; i++) {
      k = exactHeader(header[i]);
      if (k && !used[k]) {
        mapping[i] = k;
        used[k] = 1;
      }
    }
    // 第二轮：包含匹配
    for (i = 0; i < header.length; i++) {
      if (mapping[i]) continue;
      k = looseHeader(header[i], used);
      if (k && !used[k]) {
        mapping[i] = k;
        used[k] = 1;
      }
    }

    if (used.name !== 1) {
      return { payload: [], mapping: mapping, rows: 0, skippedNoName: 0, error: '没有找到「名称」列' };
    }

    var payload = [];
    var skippedNoName = 0;
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var rec = {};
      var any = false;
      for (var c = 0; c < mapping.length; c++) {
        var key = mapping[c];
        if (!key) continue;
        var v = String(row[c] == null ? '' : row[c]).trim();
        if (!v) continue;
        if (key === 'lastFollow' || key === 'createdAt') v = normalizeDate(v);
        rec[key] = v;
        any = true;
      }
      if (!any) continue;
      if (!rec.name) {
        skippedNoName++;
        continue;
      }
      payload.push(rec);
    }
    return {
      payload: payload,
      mapping: mapping,
      unmapped: header.filter(function (h, idx) {
        return normHeader(h) && !mapping[idx];
      }),
      rows: rows.length - 1,
      skippedNoName: skippedNoName,
    };
  }

  /* ---------------- 人员 / 团队 ---------------- */

  function addPersonnel(rec) {
    var name = String(rec.name || '').trim();
    if (!name) return { error: '姓名不能为空' };
    var login = String(rec.login || '').trim().toLowerCase();
    if (!login) return { error: '登录账号（名字拼音）不能为空' };
    if (state.personnel.some(function (p) { return p.login === login; })) {
      return { error: '登录账号「' + login + '」已被占用' };
    }
    var item = Object.assign({ id: uid('p'), active: true }, rec, { name: name, login: login });
    state.personnel.push(item);
    if (item.team && !state.teams.some(function (t) { return t.name === item.team; })) {
      state.teams.push({ id: uid('t'), name: item.team });
    }
    save();
    return { data: item };
  }

  function updatePersonnel(id, patch) {
    var p = state.personnel.find(function (x) { return x.id === id; });
    if (!p) return { error: '人员不存在' };
    var nextLogin = patch.login != null ? String(patch.login).trim().toLowerCase() : p.login;
    if (
      nextLogin !== p.login &&
      state.personnel.some(function (x) { return x.login === nextLogin; })
    ) {
      return { error: '登录账号「' + nextLogin + '」已被占用' };
    }
    Object.assign(p, patch, { login: nextLogin });
    // 同步已开通的账号
    var acc = state.accounts.find(function (a) { return a.personnelId === id; });
    if (acc) {
      acc.name = p.name;
      acc.login = p.login;
      acc.role = p.role;
      acc.team = p.team;
      if (acc.phone !== p.phone && p.phone) {
        acc.phone = p.phone;
        acc.pwHash = hash(p.phone);
      }
      // 当前登录者改了自己的账号，同步会话标识
      if (state.session && state.session.login !== p.login) state.session.login = p.login;
    }
    save();
    return { data: p };
  }

  function removePersonnel(id) {
    var i = state.personnel.findIndex(function (x) { return x.id === id; });
    if (i < 0) return { error: '人员不存在' };
    state.personnel.splice(i, 1);
    state.accounts = state.accounts.filter(function (a) { return a.personnelId !== id; });
    save();
    return { data: true };
  }

  function addTeam(name) {
    name = String(name || '').trim();
    if (!name) return { error: '团队名称不能为空' };
    if (state.teams.some(function (t) { return t.name === name; })) {
      return { error: '团队「' + name + '」已存在' };
    }
    var t = { id: uid('t'), name: name };
    state.teams.push(t);
    save();
    return { data: t };
  }

  function removeTeam(id) {
    var i = state.teams.findIndex(function (t) { return t.id === id; });
    if (i < 0) return { error: '团队不存在' };
    state.teams.splice(i, 1);
    save();
    return { data: true };
  }

  /* ---------------- 账号 ---------------- */

  function addAccount(rec) {
    var login = String(rec.login || '').trim().toLowerCase();
    if (!login) return { error: '登录账号不能为空' };
    if (!rec.phone) return { error: '工作手机号不能为空（作为密码）' };
    if (state.accounts.some(function (a) { return a.login === login; })) {
      return { error: '账号「' + login + '」已存在' };
    }
    var acc = {
      id: uid('a'),
      personnelId: rec.personnelId || '',
      name: String(rec.name || '').trim(),
      login: login,
      phone: String(rec.phone).trim(),
      pwHash: hash(String(rec.phone).trim()),
      role: rec.role || 'sales',
      team: rec.team || '',
      active: rec.active !== false,
      createdAt: today(),
    };
    state.accounts.push(acc);
    save();
    return { data: acc };
  }

  function updateAccount(id, patch) {
    var a = state.accounts.find(function (x) { return x.id === id; });
    if (!a) return { error: '账号不存在' };
    if (patch.phone) {
      patch.phone = String(patch.phone).trim();
      patch.pwHash = hash(patch.phone);
    }
    Object.assign(a, patch);
    save();
    return { data: a };
  }

  function removeAccount(id) {
    var i = state.accounts.findIndex(function (x) { return x.id === id; });
    if (i < 0) return { error: '账号不存在' };
    if (state.accounts[i].login === (state.session && state.session.login)) {
      return { error: '不能删除当前登录的账号' };
    }
    state.accounts.splice(i, 1);
    save();
    return { data: true };
  }

  /** 首个账号创建后即完成初始化；同时建立对应的人员档案 */
  function finishSetup(adminRec) {
    var login = String(adminRec.login || '').trim().toLowerCase();
    var name = String(adminRec.name || '').trim();

    // 种子人员里可能已经存在同名/同账号的人（例如 CRM 业务员名单里的本人），直接接管
    var existing =
      state.personnel.find(function (p) {
        return p.login === login;
      }) ||
      state.personnel.find(function (p) {
        return p.name === name;
      });

    var person;
    if (existing) {
      existing.role = 'admin';
      existing.active = true;
      if (adminRec.phone) existing.phone = String(adminRec.phone).trim();
      existing.note = '由初始化向导指定为主管理员';
      // 账号沿用一条自有的登录名，避免与其它记录冲突
      if (state.personnel.some(function (p) {
        return p !== existing && p.login === login;
      })) {
        return { error: '登录账号「' + login + '」已被其它人员占用，请换一个' };
      }
      existing.login = login;
      person = { data: existing };
      save();
    } else {
      person = addPersonnel({
        name: name,
        login: login,
        role: 'admin',
        team: adminRec.team || '',
        manager: '',
        phone: adminRec.phone,
        note: '首位主管理员，由初始化向导创建',
      });
      if (person.error) return person;
    }

    var r = addAccount(
      Object.assign({ role: 'admin', personnelId: person.data.id, name: name }, adminRec, { login: login })
    );
    if (r.error) {
      if (!existing) removePersonnel(person.data.id);
      return r;
    }
    state.meta.initialized = true;
    save();
    return r;
  }

  /* ---------------- 待办任务 ---------------- */

  function addTask(rec) {
    var content = String(rec.content || '').trim();
    if (!content) return { error: '事项内容不能为空' };
    var t = {
      id: uid('k'),
      date: normalizeDate(rec.date) || today(),
      customer: String(rec.customer || '').trim(),
      content: content,
      note: String(rec.note || '').trim(),
      owner: rec.owner || (currentUser() && currentUser().login) || '',
      creator: (currentUser() && currentUser().login) || '',
      priority: rec.priority || '普通',
      orderRef: rec.orderRef || '',
      done: false,
      doneAt: null,
      createdAt: nowStamp(),
    };
    state.tasks.push(t);
    save();
    return { data: t };
  }

  function updateTask(id, patch) {
    var t = state.tasks.find(function (x) { return x.id === id; });
    if (!t) return { error: '任务不存在' };
    Object.assign(t, patch);
    save();
    return { data: t };
  }

  /** 切换完成状态；完成后记录点击时间（导出用） */
  function toggleTask(id, force) {
    var t = state.tasks.find(function (x) { return x.id === id; });
    if (!t) return { error: '任务不存在' };
    var next = force == null ? !t.done : !!force;
    t.done = next;
    t.doneAt = next ? nowStamp() : null;
    save();
    return { data: t };
  }

  function removeTask(id) {
    var i = state.tasks.findIndex(function (x) { return x.id === id; });
    if (i < 0) return { error: '任务不存在' };
    state.tasks.splice(i, 1);
    save();
    return { data: true };
  }

  /**
   * 导出用的分组：按「完成日期 + 客户」合并，同一客户的多条事项换行。
   * 依据 SOP：时间为点击完成的时间。
   */
  function groupDoneTasks(opts) {
    opts = opts || {};
    var from = opts.from || '';
    var to = opts.to || '';
    var owner = opts.owner || '';
    var map = {};
    var order = [];
    state.tasks
      .filter(function (t) { return t.done && t.doneAt; })
      .forEach(function (t) {
        var day = dateOnly(t.doneAt);
        if (from && day < from) return;
        if (to && day > to) return;
        if (owner && t.owner !== owner) return;
        var key = day + '||' + (t.customer || '（未填客户）');
        if (!map[key]) {
          map[key] = {
            day: day,
            customer: t.customer || '（未填客户）',
            items: [],
            note: '',
            owners: [],
            lastAt: t.doneAt,
          };
          order.push(key);
        }
        var g = map[key];
        g.items.push(t.content);
        if (t.note && !g.note) g.note = t.note;
        var ownerName = nameOfLogin(t.owner);
        if (ownerName && g.owners.indexOf(ownerName) === -1) g.owners.push(ownerName);
        if (t.doneAt > g.lastAt) g.lastAt = t.doneAt;
      });
    order.sort(function (a, b) {
      return a < b ? 1 : a > b ? -1 : 0;
    });
    return order.map(function (k) {
      var g = map[k];
      return {
        day: g.day,
        customer: g.customer,
        content: g.items.join('\n'),
        count: g.items.length,
        owner: g.owners.join('、'),
        note: g.note,
        doneAt: g.lastAt,
      };
    });
  }

  function nameOfLogin(login) {
    if (!login) return '';
    var p = state.personnel.find(function (x) { return x.login === login; });
    if (p) return p.name;
    var a = state.accounts.find(function (x) { return x.login === login; });
    return a ? a.name : login;
  }

  /* ---------------- 全量备份 ---------------- */

  function exportAll() {
    return JSON.stringify(state, null, 2);
  }

  function importAll(json) {
    var parsed = typeof json === 'string' ? JSON.parse(json) : json;
    if (!parsed || typeof parsed !== 'object' || !Array.isArray(parsed.tasks)) {
      return { error: '文件格式不正确，应为平台导出的备份 JSON' };
    }
    var sessionLogin = state.session && state.session.login;
    state = Object.assign(blankState(), parsed);
    if (sessionLogin && state.accounts.some(function (a) { return a.login === sessionLogin; })) {
      state.session = { login: sessionLogin, at: nowStamp() };
    } else {
      state.session = null;
    }
    save();
    return { data: true };
  }

  function wipe() {
    localStorage.removeItem(KEY);
    state = blankState();
    save();
  }

  root.Store = {
    KEY: KEY,
    ROLES: ROLES,
    TASK_STATUS: TASK_STATUS,
    DEFAULT_NAV: DEFAULT_NAV,
    get state() { return state; },
    load: load,
    save: save,
    seedIfEmpty: seedIfEmpty,
    uid: uid,
    today: today,
    nowStamp: nowStamp,
    normalizeDate: normalizeDate,
    dateOnly: dateOnly,
    hash: hash,
    currentUser: currentUser,
    signIn: signIn,
    signOut: signOut,
    roleOf: roleOf,
    isAdmin: isAdmin,
    canAssign: canAssign,
    assignableTargets: assignableTargets,
    addCustomer: addCustomer,
    updateCustomer: updateCustomer,
    removeCustomer: removeCustomer,
    importCustomers: importCustomers,
    customerKey: customerKey,
    CUSTOMER_HEADERS: CUSTOMER_HEADERS,
    mapCustomerHeader: mapCustomerHeader,
    mapCustomerRows: mapCustomerRows,
    addPersonnel: addPersonnel,
    updatePersonnel: updatePersonnel,
    removePersonnel: removePersonnel,
    addTeam: addTeam,
    removeTeam: removeTeam,
    addAccount: addAccount,
    updateAccount: updateAccount,
    removeAccount: removeAccount,
    finishSetup: finishSetup,
    addTask: addTask,
    updateTask: updateTask,
    toggleTask: toggleTask,
    removeTask: removeTask,
    groupDoneTasks: groupDoneTasks,
    nameOfLogin: nameOfLogin,
    exportAll: exportAll,
    importAll: importAll,
    wipe: wipe,
  };
})(typeof window !== 'undefined' ? window : globalThis);
