/*!
 * app.js — 销售&跟单工作平台 界面与交互
 * 依据《销售跟单平台SOP》实现五个界面。
 */
(function () {
  'use strict';

  var S = window.Store;
  var XLSX = window.XlsxLite;
  var READ = window.XlsxRead;

  /* ============================ 界面状态 ============================ */

  var ui = {
    view: 'todo',
    openNav: { data: true },
    search: '',
    // 待办
    sortBy: 'date',
    showDone: false,
    fCustomer: '',
    fFrom: '',
    fTo: '',
    fOwner: '',
    todoPage: 1,
    // 今日抽屉
    todayOpen: false,
    // 客户表
    custPage: 1,
    custSize: 50,
    custSource: '',
    custLevel: '',
    custSales: '',
    // 导出
    exFrom: '',
    exTo: '',
    exOwner: '',
    exMode: 'group',
    // 人员
    personRole: '',
    // 初始化向导阶段
    setupStage: 'login',
    pendingSetup: null,
  };

  var PAGE_SIZE = 60;

  /* ============================ 工具 ============================ */

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function el(id) {
    return document.getElementById(id);
  }

  function toast(msg, kind) {
    var host = el('toastHost');
    var node = document.createElement('div');
    node.className = 'toast ' + (kind || '');
    node.innerHTML = esc(msg);
    host.appendChild(node);
    setTimeout(function () {
      node.style.transition = 'opacity .25s, transform .25s';
      node.style.opacity = '0';
      node.style.transform = 'translateY(6px)';
      setTimeout(function () {
        if (node.parentNode) node.parentNode.removeChild(node);
      }, 260);
    }, 3400);
  }

  function roleLabel(r) {
    return (S.ROLES[r] || {}).label || r || '';
  }

  function initials(name) {
    var n = String(name || '?').trim();
    return n.slice(-2);
  }

  function monthStart() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-01';
  }

  function weekStart() {
    var d = new Date();
    var day = (d.getDay() + 6) % 7; // 周一为一周起点
    d.setDate(d.getDate() - day);
    return (
      d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0')
    );
  }

  function downloadText(text, filename, mime) {
    var blob = new Blob([text], { type: mime || 'application/json;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 2000);
  }

  function openModal(title, bodyHtml, footHtml, wide) {
    closeModal();
    var mask = document.createElement('div');
    mask.className = 'modal-mask';
    mask.id = 'modalMask';
    mask.innerHTML =
      '<div class="modal' + (wide ? ' wide' : '') + '">' +
      '<div class="modal-head"><h3>' + esc(title) + '</h3>' +
      '<button class="btn-icon" data-action="close-modal" title="关闭">✕</button></div>' +
      '<div class="modal-body">' + bodyHtml + '</div>' +
      (footHtml ? '<div class="modal-foot">' + footHtml + '</div>' : '') +
      '</div>';
    mask.addEventListener('click', function (e) {
      if (e.target === mask) closeModal();
    });
    document.body.appendChild(mask);
  }

  function closeModal() {
    var m = el('modalMask');
    if (m && m.parentNode) m.parentNode.removeChild(m);
  }

  function uniqueSorted(arr) {
    var seen = {};
    var out = [];
    arr.forEach(function (v) {
      v = String(v == null ? '' : v).trim();
      if (!v || seen[v]) return;
      seen[v] = 1;
      out.push(v);
    });
    return out.sort(function (a, b) {
      return a.localeCompare(b, 'zh-Hans-CN');
    });
  }

  /* ============================ 渲染入口 ============================ */

  function render() {
    var user = S.currentUser();
    if (!S.state.meta.initialized || !user) {
      renderAuth();
      return;
    }
    el('root').innerHTML =
      '<div class="app">' +
      renderHeader(user) +
      '<div class="app-body">' +
      renderSidebar(user) +
      '<div class="main" id="main">' +
      renderMain(user) +
      '</div>' +
      '</div>' +
      '</div>' +
      (ui.todayOpen ? renderTodayDrawer(user) : '');
  }

  /* ============================ 登录 / 初始化 ============================ */

  function authShell(inner) {
    return '<div class="auth-wrap"><div class="auth-card">' + inner + '</div></div>';
  }

  function renderAuth() {
    var initialized = S.state.meta.initialized;
    var stage = ui.setupStage; // 'setup' | 'pending' | 'login'

    // —— 已完成初始化：显示登录 ——
    if (initialized) {
      if (stage === 'pending' && ui.pendingSetup) {
        var p = ui.pendingSetup;
        el('root').innerHTML = authShell(
          '<img class="auth-logo" src="assets/logo.png" alt="QUICK SHOW 捷展">' +
            '<h1>账号已创建</h1>' +
            '<p class="auth-sub">请用下面的信息登录，密码为你的工作手机号。</p>' +
            '<div class="alert alert-ok">登录账号：<b>' + esc(p.login) + '</b><br>密码：<b>' + esc(p.phone) + '</b></div>' +
            '<button class="btn btn-primary btn-block" data-action="go-login">前往登录</button>'
        );
        return;
      }

      el('root').innerHTML = authShell(
        '<img class="auth-logo" src="assets/logo.png" alt="QUICK SHOW 捷展">' +
          '<h1>销售&amp;跟单工作平台</h1>' +
          '<p class="auth-sub">QUICK SHOW 捷展 · 潮牌供应链</p>' +
          '<form id="loginForm">' +
          '<div class="field"><label>账号</label>' +
          '<input class="input" id="loginName" autocomplete="username" placeholder="姓名的拼音，例如 zhangwei" required></div>' +
          '<div class="field"><label>密码<span class="hint">即你的工作手机号</span></label>' +
          '<input class="input" id="loginPwd" type="password" autocomplete="current-password" placeholder="请输入工作手机号" required></div>' +
          '<div id="loginErr"></div>' +
          '<button class="btn btn-primary btn-block" type="submit">登 录</button>' +
          '</form>' +
          '<div class="auth-note">账号由主管理员在「数据源表格 → 人员架构表 / 账号登录」中开通。<br>忘记密码请联系主管理员。</div>'
      );
      return;
    }

    // —— 尚未初始化：创建主管理员 ——
    el('root').innerHTML = authShell(
      '<img class="auth-logo" src="assets/logo.png" alt="QUICK SHOW 捷展">' +
        '<h1>初始化平台</h1>' +
        '<p class="auth-sub">首次使用，请先创建主管理员账号。</p>' +
        (stage === 'error' && ui.setupError
          ? '<div class="alert alert-danger">' + esc(ui.setupError) + '</div>'
          : '') +
        '<form id="setupForm">' +
        '<div class="grid-2">' +
        '<div class="field"><label>姓名</label><input class="input" id="suName" placeholder="例如 彭培强" required></div>' +
        '<div class="field"><label>登录账号<span class="hint">名字拼音</span></label>' +
        '<input class="input" id="suLogin" placeholder="例如 pengpeiqiang" required></div>' +
        '</div>' +
        '<div class="field"><label>工作手机号<span class="hint">将作为登录密码</span></label>' +
        '<input class="input" id="suPhone" placeholder="例如 13800138000" required></div>' +
        '<button class="btn btn-primary btn-block" type="submit">创建并进入平台</button>' +
        '</form>' +
        '<div class="auth-note">平台已内置 <b>' + S.state.customers.length + '</b> 位客户与 <b>' +
        S.state.personnel.length +
        '</b> 位业务员资料（来自 CRM 客户表）。<br>创建后可继续在「数据源表格」中补充团队、主管与跟单人员。</div>'
    );
  }

  /* ============================ 顶栏 ============================ */

  function renderHeader(user) {
    return (
      '<header class="app-header">' +
      '<div class="brand">' +
      '<img class="brand-logo" src="assets/logo.png" alt="QUICK SHOW 捷展">' +
      '<div class="brand-divider"></div>' +
      '<div class="brand-text">' +
      '<div class="brand-title">销售&amp;跟单工作平台</div>' +
      '<div class="brand-sub">QUICK SHOW · 潮牌供应链</div>' +
      '</div></div>' +
      '<div class="header-spacer"></div>' +
      '<div class="header-tools">' +
      '<div class="search-box' + (ui.search ? ' has-value' : '') + '">' +
      '<input class="input" id="globalSearch" placeholder="搜索事项 / 客户 / 备注" value="' + esc(ui.search) + '">' +
      '<button class="search-clear" data-action="clear-search" title="清空">✕</button>' +
      '</div>' +
      '<button class="btn" data-action="open-today">今日任务' +
      (todayCount() ? '<span class="nav-badge">' + todayCount() + '</span>' : '') +
      '</button>' +
      '<div class="user-chip" data-action="toggle-user-menu" title="' + esc(user.name) + ' · ' + esc(roleLabel(user.role)) + '">' +
      '<div class="avatar">' + esc(initials(user.name)) + '</div>' +
      '<div class="user-meta"><div class="n">' + esc(user.name) + '</div>' +
      '<div class="r">' + esc(roleLabel(user.role)) + '</div></div>' +
      '</div>' +
      '<button class="btn btn-ghost btn-sm" data-action="sign-out">退出</button>' +
      '</div>' +
      '</header>'
    );
  }

  /* ============================ 侧边导航 ============================ */

  function renderSidebar(user) {
    var nav = S.state.navOrder;
    var html = '<aside class="sidebar">';
    html += '<div class="sidebar-hint">≡ 按住拖动可调整顺序</div>';

    nav.forEach(function (item, i) {
      var children = item.children || [];
      var isOpen = !!ui.openNav[item.id];
      var selfActive = ui.view === item.id;
      var childActive = children.some(function (c) {
        return c.id === ui.view;
      });
      var badge = '';
      if (item.id === 'todo') {
        var n = openTaskCount();
        if (n) badge = '<span class="nav-badge">' + n + '</span>';
      }

      html +=
        '<div class="nav-item' + (selfActive ? ' active' : '') + (isOpen ? ' open' : '') +
        '" draggable="true" data-nav="' + esc(item.id) + '" data-index="' + i + '">' +
        '<span class="nav-grip">⋮⋮</span>' +
        '<span class="nav-icon">' + esc(item.icon || '•') + '</span>' +
        '<span class="nav-label">' + esc(item.label) + '</span>' +
        badge +
        (children.length ? '<span class="nav-caret">▶</span>' : '') +
        '</div>';

      if (children.length) {
        html +=
          '<div class="nav-children' + (isOpen || childActive ? ' open' : '') + '">' +
          children
            .map(function (c) {
              return (
                '<div class="nav-child' + (ui.view === c.id ? ' active' : '') +
                '" data-nav-child="' + esc(c.id) + '">' + esc(c.label) + '</div>'
              );
            })
            .join('') +
          '</div>';
      }
    });

    html += '</aside>';
    return html;
  }

  /* ============================ 主内容路由 ============================ */

  function renderMain(user) {
    switch (ui.view) {
      case 'todo':
        return viewTodo(user);
      case 'orders':
        return viewOrders(user);
      case 'assign':
        return viewAssign(user);
      case 'export':
        return viewExport(user);
      case 'data-customers':
        return viewCustomers(user);
      case 'data-personnel':
        return viewPersonnel(user);
      case 'data-accounts':
        return viewAccounts(user);
      default:
        return viewTodo(user);
    }
  }

  /* ============================ 界面1：日常待办事项 ============================ */

  function openTaskCount() {
    var me = S.currentUser();
    if (!me) return 0;
    return S.state.tasks.filter(function (t) {
      return !t.done && t.owner === me.login;
    }).length;
  }

  function todayCount() {
    var me = S.currentUser();
    if (!me) return 0;
    var d = S.today();
    return S.state.tasks.filter(function (t) {
      return !t.done && t.owner === me.login && S.dateOnly(t.date) === d;
    }).length;
  }

  /**
   * 按当前筛选/排序返回待办列表。
   * 依 SOP：点完成后自动隐藏，但一旦搜索或按客户/日期筛选，隐藏的也要能重新出现。
   */
  function filteredTasks() {
    var q = ui.search.trim().toLowerCase();
    var hasFilter = !!(q || ui.fCustomer || ui.fFrom || ui.fTo || ui.fOwner);
    var list = S.state.tasks.filter(function (t) {
      if (t.done && !ui.showDone && !hasFilter) return false;
      if (ui.fCustomer && (t.customer || '') !== ui.fCustomer) return false;
      if (ui.fOwner && t.owner !== ui.fOwner) return false;
      var d = S.dateOnly(t.date);
      if (ui.fFrom && d && d < ui.fFrom) return false;
      if (ui.fTo && d && d > ui.fTo) return false;
      if (q) {
        var hay = [t.content, t.customer, t.note, S.nameOfLogin(t.owner), t.date]
          .join(' ')
          .toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });

    if (ui.sortBy === 'customer') {
      list.sort(function (a, b) {
        var c = String(a.customer || '').localeCompare(String(b.customer || ''), 'zh-Hans-CN');
        if (c) return c;
        return String(a.date).localeCompare(String(b.date));
      });
    } else {
      list.sort(function (a, b) {
        var c = String(b.date).localeCompare(String(a.date));
        if (c) return c;
        return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
      });
    }
    return list;
  }

  function viewTodo(user) {
    var list = filteredTasks();
    var total = list.length;
    var pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    if (ui.todoPage > pages) ui.todoPage = pages;
    var page = list.slice((ui.todoPage - 1) * PAGE_SIZE, ui.todoPage * PAGE_SIZE);
    var allTasks = S.state.tasks;
    var mine = allTasks.filter(function (t) {
      return !t.done && t.owner === user.login;
    }).length;
    var todayN = todayCount();
    var doneToday = allTasks.filter(function (t) {
      return t.done && S.dateOnly(t.doneAt) === S.today() && t.owner === user.login;
    }).length;
    var overdue = allTasks.filter(function (t) {
      return !t.done && t.owner === user.login && S.dateOnly(t.date) && S.dateOnly(t.date) < S.today();
    }).length;

    var canPickOwner = S.assignableTargets(user);

    var html = '';

    html +=
      '<div class="page-head"><div>' +
      '<div class="page-title">日常待办事项</div>' +
      '<div class="page-desc">按「日期 - 客户 - 事项 - 是否完成 - 备注」记录；可直接在表格中录入与修改。</div>' +
      '</div><div class="page-actions">' +
      '<button class="btn" data-action="export-todo">导出当前列表</button>' +
      '<button class="btn btn-primary" data-action="focus-add">＋ 新增任务</button>' +
      '</div></div>';

    html +=
      '<div class="stat-row">' +
      stat('我的未完成', mine, '当前账号名下待办', '') +
      stat('今日待办', todayN, '计划日期为今天', 'is-warn') +
      stat('今日已完成', doneToday, '今天点击完成的条数', 'is-ok') +
      stat('已逾期', overdue, '计划日期早于今天', overdue ? 'is-warn' : '') +
      '</div>';

    // 工具条
    html +=
      '<div class="card"><div class="card-head">' +
      '<div class="seg">' +
      '<button class="' + (ui.sortBy === 'date' ? 'active' : '') + '" data-action="sort" data-value="date">按日期排列</button>' +
      '<button class="' + (ui.sortBy === 'customer' ? 'active' : '') + '" data-action="sort" data-value="customer">按客户排列</button>' +
      '</div>' +
      '<label class="inline-filter"><input type="checkbox" class="check" data-action="toggle-showdone"' +
      (ui.showDone ? ' checked' : '') + '> 显示已完成</label>' +
      '<div class="card-head-actions">' +
      '<select class="select select-sm" data-action="filter-owner">' +
      '<option value="">全部负责人</option>' +
      ownerOptions(ui.fOwner) +
      '</select>' +
      '<select class="select select-sm" data-action="filter-customer">' +
      '<option value="">全部客户</option>' +
      customerOptions(ui.fCustomer) +
      '</select>' +
      '<input class="input input-sm" type="date" data-action="filter-from" value="' + esc(ui.fFrom) + '" title="起始日期">' +
      '<span style="color:var(--ink-4)">–</span>' +
      '<input class="input input-sm" type="date" data-action="filter-to" value="' + esc(ui.fTo) + '" title="结束日期">' +
      '<button class="btn btn-sm" data-action="reset-filters">重置</button>' +
      '</div></div>';

    html += '<div class="card-body tight"><div class="table-scroll"><table class="tbl">';
    html +=
      '<thead><tr>' +
      '<th style="width:46px">完成</th>' +
      '<th style="width:132px">日期</th>' +
      '<th style="width:168px">客户</th>' +
      '<th>事项</th>' +
      '<th style="width:150px">备注</th>' +
      '<th style="width:120px">负责人</th>' +
      '<th style="width:74px"></th>' +
      '</tr></thead><tbody>';

    // 快捷录入行
    html +=
      '<tr style="background:var(--bg-softer)">' +
      '<td></td>' +
      '<td><input class="cell-input is-date" type="date" data-new="date" value="' + S.today() + '"></td>' +
      '<td><input class="cell-input" data-new="customer" list="customerList" placeholder="客户名称"></td>' +
      '<td><input class="cell-input" data-new="content" id="quickContent" placeholder="输入事项内容，回车即可新增"></td>' +
      '<td><input class="cell-input" data-new="note" placeholder="备注"></td>' +
      '<td><select class="cell-input" data-new="owner">' +
      ownerOptions(user.login, canPickOwner) +
      '</select></td>' +
      '<td class="actions"><button class="btn btn-primary btn-xs" data-action="add-task">新增</button></td>' +
      '</tr>';

    if (!page.length) {
      html +=
        '<tr><td colspan="7"><div class="empty"><div class="empty-mark">☑</div>' +
        '<h3>没有符合条件的待办</h3><p>' +
        (ui.showDone ? '换个筛选条件看看。' : '已完成的事项会自动隐藏，勾选上方「显示已完成」或调整客户/日期筛选即可找回。') +
        '</p></div></td></tr>';
    }

    page.forEach(function (t) {
      html +=
        '<tr data-id="' + esc(t.id) + '" class="' + (t.done ? 'is-done' : '') + '">' +
        '<td><input type="checkbox" class="check" data-action="toggle-task"' + (t.done ? ' checked' : '') + '></td>' +
        '<td><input class="cell-input is-date" type="date" data-field="date" value="' + esc(S.dateOnly(t.date)) + '"></td>' +
        '<td><input class="cell-input" data-field="customer" list="customerList" value="' + esc(t.customer) + '" placeholder="客户"></td>' +
        '<td><input class="cell-input task-content" data-field="content" value="' + esc(t.content) + '"></td>' +
        '<td><input class="cell-input" data-field="note" value="' + esc(t.note) + '" placeholder="—"></td>' +
        '<td><select class="cell-input" data-field="owner">' +
        ownerOptions(t.owner, canPickOwner) +
        '</select></td>' +
        '<td class="actions"><button class="btn-icon" data-action="del-task" title="删除">✕</button></td>' +
        '</tr>';
    });

    html += '</tbody></table></div>';

    if (total > PAGE_SIZE) {
      html +=
        '<div class="pager"><span>共 ' + total + ' 条，第 ' + ui.todoPage + ' / ' + pages + ' 页</span>' +
        '<span class="spacer"></span>' +
        '<button class="btn btn-sm" data-action="todo-prev"' + (ui.todoPage <= 1 ? ' disabled' : '') + '>上一页</button>' +
        '<button class="btn btn-sm" data-action="todo-next"' + (ui.todoPage >= pages ? ' disabled' : '') + '>下一页</button>' +
        '</div>';
    }

    html += '</div></div>';
    html += '<datalist id="customerList">' + customerOptions('') + '</datalist>';

    return html;
  }

  function stat(label, value, foot, cls) {
    return (
      '<div class="stat ' + (cls || '') + '">' +
      '<div class="stat-label">' + esc(label) + '</div>' +
      '<div class="stat-value">' + esc(value) + '</div>' +
      '<div class="stat-foot">' + esc(foot) + '</div>' +
      '</div>'
    );
  }

  function ownerOptions(selected, restrictList) {
    var people = S.state.personnel.filter(function (p) {
      return p.active;
    });
    if (restrictList) {
      var allowed = {};
      restrictList.forEach(function (p) {
        allowed[p.login] = 1;
      });
      people = people.filter(function (p) {
        return allowed[p.login];
      });
    }
    var html = people
      .map(function (p) {
        return (
          '<option value="' + esc(p.login) + '"' + (p.login === selected ? ' selected' : '') + '>' +
          esc(p.name) + '</option>'
        );
      })
      .join('');
    if (selected && !people.some(function (p) {
      return p.login === selected;
    })) {
      html =
        '<option value="' + esc(selected) + '" selected>' + esc(S.nameOfLogin(selected)) + '</option>' + html;
    }
    return html;
  }

  function customerOptions(selected) {
    var names = S.state.customers
      .map(function (c) {
        return c.name;
      })
      .filter(Boolean);
    return uniqueSorted(names)
      .slice(0, 3000)
      .map(function (n) {
        return '<option value="' + esc(n) + '"' + (n === selected ? ' selected' : '') + '></option>';
      })
      .join('');
  }

  /* ============================ 今日任务抽屉 ============================ */

  function renderTodayDrawer(user) {
    var d = S.today();
    var mine = S.state.tasks.filter(function (t) {
      return t.owner === user.login && S.dateOnly(t.date) === d;
    });
    var undone = mine.filter(function (t) {
      return !t.done;
    });
    var done = mine.filter(function (t) {
      return t.done;
    });
    var overdue = S.state.tasks.filter(function (t) {
      return t.owner === user.login && !t.done && S.dateOnly(t.date) && S.dateOnly(t.date) < d;
    });

    function item(t) {
      return (
        '<div class="today-item"><input type="checkbox" class="check" data-action="toggle-task" data-id="' +
        esc(t.id) + '"' + (t.done ? ' checked' : '') + '>' +
        '<div class="body"><div class="t1">' +
        (t.customer ? '<b>' + esc(t.customer) + '</b> · ' : '') + esc(t.content) + '</div>' +
        '<div class="t2">' + esc(t.date) + (t.note ? ' · ' + esc(t.note) : '') +
        (t.done && t.doneAt ? ' · 完成于 ' + esc(t.doneAt.slice(11)) : '') + '</div></div></div>'
      );
    }

    var html =
      '<div class="drawer-mask" data-action="close-today"><div class="drawer" data-stop="1">' +
      '<div class="drawer-head"><h3>' + d + ' 今日任务</h3>' +
      '<button class="btn-icon" data-action="close-today">✕</button></div>' +
      '<div class="drawer-body">';

    html +=
      '<div class="drawer-section-title">未完成 <span class="count">' + undone.length + '</span></div>';
    html += undone.length ? undone.map(item).join('') : '<div class="empty" style="padding:22px"><p>今日没有未完成事项 🎉</p></div>';

    if (overdue.length) {
      html +=
        '<div class="drawer-section-title">已逾期 <span class="count">' + overdue.length + '</span></div>' +
        overdue
          .slice(0, 30)
          .map(function (t) {
            return (
              '<div class="today-item"><input type="checkbox" class="check" data-action="toggle-task" data-id="' +
              esc(t.id) + '"><div class="body"><div class="t1">' +
              (t.customer ? '<b>' + esc(t.customer) + '</b> · ' : '') + esc(t.content) + '</div>' +
              '<div class="t2">原计划 ' + esc(t.date) + '</div></div></div>'
            );
          })
          .join('');
    }

    html +=
      '<div class="drawer-section-title">已完成 <span class="count">' + done.length + '</span></div>';
    html += done.length ? done.map(item).join('') : '<div class="empty" style="padding:22px"><p>今天还没有完成的事项</p></div>';

    html += '</div></div></div>';
    return html;
  }

  /* ============================ 界面2：大货订单跟进 ============================ */

  function viewOrders() {
    return (
      '<div class="page-head"><div>' +
      '<div class="page-title">大货订单跟进</div>' +
      '<div class="page-desc">按 SOP 约定，本模块暂缓建设，先保留位置。</div>' +
      '</div></div>' +
      '<div class="card"><div class="card-body">' +
      '<div class="empty"><div class="empty-mark">▤</div>' +
      '<h3>模块待建设</h3>' +
      '<p>SOP 中此处标注为「先放空」。数据结构已预留，后续可直接接入大货订单的进度节点、交期与跟单记录，无需改动其它模块。</p>' +
      '</div></div></div>'
    );
  }

  /* ============================ 界面3：任务指派 ============================ */

  function viewAssign(user) {
    var targets = S.assignableTargets(user);
    var myAssignments = S.state.tasks
      .filter(function (t) {
        return t.creator === user.login;
      })
      .sort(function (a, b) {
        return String(b.createdAt || '').localeCompare(String(a.createdAt || ''));
      })
      .slice(0, 40);

    var html = '';

    html +=
      '<div class="page-head"><div>' +
      '<div class="page-title">任务指派</div>' +
      '<div class="page-desc">按「主管理员 → 团队主管 → 销售 → 跟单」的顺序向下指派，指派结果会直接出现在对方的「日常待办事项」中。</div>' +
      '</div></div>';

    // 指派表单
    html +=
      '<div class="card"><div class="card-head"><div class="card-title">发起指派</div>' +
      '<div class="card-sub">当前身份：' + esc(roleLabel(user.role)) + '</div></div>' +
      '<div class="card-body"><form id="assignForm">' +
      '<div class="grid-3">' +
      '<div class="field"><label>指派给</label><select class="select" id="asOwner" required>' +
      targets
        .map(function (p) {
          return (
            '<option value="' + esc(p.login) + '">' + esc(p.name) + '（' + esc(roleLabel(p.role)) +
            (p.team ? ' · ' + esc(p.team) : '') + '）</option>'
          );
        })
        .join('') +
      '</select></div>' +
      '<div class="field"><label>计划日期</label><input class="input" type="date" id="asDate" value="' + S.today() + '"></div>' +
      '<div class="field"><label>优先级</label><select class="select" id="asPriority">' +
      '<option>普通</option><option>紧急</option><option>重要</option></select></div>' +
      '</div>' +
      '<div class="grid-2">' +
      '<div class="field"><label>客户</label><input class="input" id="asCustomer" list="customerListAssign" placeholder="输入或选择客户"></div>' +
      '<div class="field"><label>备注</label><input class="input" id="asNote" placeholder="补充说明"></div>' +
      '</div>' +
      '<div class="field"><label>事项内容</label>' +
      '<textarea class="textarea" id="asContent" placeholder="例如：跟进 9 月新品报价，周五前给出方案" required></textarea></div>' +
      '<button class="btn btn-primary" type="submit">发送任务</button>' +
      '</form></div></div>';

    html += '<datalist id="customerListAssign">' + customerOptions('') + '</datalist>';

    // 权限矩阵
    html +=
      '<div class="card"><div class="card-head"><div class="card-title">指派权限矩阵</div>' +
      '<div class="card-sub">规则来源：SOP 中「主管理员-团队主管-销售-跟单」顺序，且跟单也可给销售指派</div></div>' +
      '<div class="card-body">' +
      buildMatrix() +
      '</div></div>';

    // 组织架构
    html +=
      '<div class="card"><div class="card-head"><div class="card-title">组织架构</div>' +
      '<div class="card-sub">在「数据源表格 → 人员架构表」中维护</div></div>' +
      '<div class="card-body"><div class="tree">' + buildTree() + '</div></div></div>';

    // 我发出的指派
    html +=
      '<div class="card"><div class="card-head"><div class="card-title">我发出的指派</div>' +
      '<div class="card-sub">最近 ' + myAssignments.length + ' 条</div></div>' +
      '<div class="card-body tight"><div class="table-scroll"><table class="tbl"><thead><tr>' +
      '<th style="width:132px">日期</th><th style="width:170px">客户</th><th>事项</th>' +
      '<th style="width:110px">指派给</th><th style="width:86px">状态</th>' +
      '</tr></thead><tbody>';

    if (!myAssignments.length) {
      html +=
        '<tr><td colspan="5"><div class="empty" style="padding:32px"><p>还没有发出过指派。</p></div></td></tr>';
    }
    myAssignments.forEach(function (t) {
      html +=
        '<tr><td class="nowrap">' + esc(t.date) + '</td>' +
        '<td>' + esc(t.customer || '—') + '</td>' +
        '<td class="task-content">' + esc(t.content) + '</td>' +
        '<td>' + esc(S.nameOfLogin(t.owner)) + '</td>' +
        '<td>' + (t.done
          ? '<span class="tag tag-ok">已完成</span>'
          : '<span class="tag tag-muted">进行中</span>') + '</td></tr>';
    });

    html += '</tbody></table></div></div></div>';
    return html;
  }

  function buildMatrix() {
    var roles = ['admin', 'manager', 'sales', 'followup'];
    var html =
      '<table class="matrix"><thead><tr><th class="rowhead">发起人 \\ 可指派给</th>' +
      roles
        .map(function (r) {
          return '<th>' + esc(roleLabel(r)) + '</th>';
        })
        .join('') +
      '</tr></thead><tbody>';

    roles.forEach(function (from) {
      html += '<tr><td class="rowhead">' + esc(roleLabel(from)) + '</td>';
      roles.forEach(function (to) {
        var fake = { role: from, team: '销售部', login: 'x' };
        var fakeTarget = { role: to, team: '销售部', login: 'y' };
        var yes = S.canAssign(fake, fakeTarget);
        html += '<td class="' + (yes ? 'yes' : 'no') + '">' + (yes ? '✓' : '—') + '</td>';
      });
      html += '</tr>';
    });

    html += '</tbody></table>';
    html +=
      '<div class="alert alert-info" style="margin-top:14px;margin-bottom:0">' +
      '说明：主管理员可指派给任何人；团队主管限本团队成员；销售与跟单之间可互相指派（SOP 明确「跟单也可以给销售指派任务」）。' +
      '</div>';
    return html;
  }

  function buildTree() {
    var people = S.state.personnel.filter(function (p) {
      return p.active;
    });
    if (!people.length) return '<div class="empty" style="padding:24px"><p>还没有人员资料。</p></div>';

    var teams = {};
    people.forEach(function (p) {
      var t = p.team || '未分组';
      (teams[t] = teams[t] || []).push(p);
    });

    var admins = people.filter(function (p) {
      return p.role === 'admin';
    });

    var html = '<div class="tree-root">';
    html +=
      '<div class="tree-row"><span class="tag tag-brand">主管理员</span>' +
      (admins.length
        ? admins.map(function (a) { return '<span class="tree-name">' + esc(a.name) + '</span>'; }).join('、')
        : '<span style="color:var(--ink-4)">尚未指定</span>') +
      '</div>';

    Object.keys(teams)
      .sort()
      .forEach(function (tname) {
        var members = teams[tname];
        var leaders = members.filter(function (m) {
          return m.role === 'manager';
        });
        html += '<div class="tree-node">';
        html +=
          '<div class="tree-row"><span class="tag">' + esc(tname) + '</span>' +
          (leaders.length
            ? '<span class="tree-name">' + leaders.map(function (l) { return esc(l.name); }).join('、') + '</span><span class="card-sub">团队主管</span>'
            : '<span style="color:var(--ink-4)">未设主管</span>') +
          '</div>';

        ['sales', 'followup'].forEach(function (role) {
          var group = members.filter(function (m) {
            return m.role === role;
          });
          if (!group.length) return;
          html +=
            '<div class="tree-node"><div class="tree-row"><span class="tag tag-muted">' + esc(roleLabel(role)) +
            '</span><span>' +
            group
              .map(function (g) {
                return '<span class="tree-name" style="margin-right:12px">' + esc(g.name) + '</span>';
              })
              .join('') +
            '</span></div></div>';
        });

        html += '</div>';
      });

    html += '</div>';
    return html;
  }

  /* ============================ 界面4：每日跟进记录导出 ============================ */

  function viewExport(user) {
    var groups = S.groupDoneTasks({ from: ui.exFrom, to: ui.exTo, owner: ui.exOwner });
    var detail = S.state.tasks
      .filter(function (t) {
        if (!t.done || !t.doneAt) return false;
        var d = S.dateOnly(t.doneAt);
        if (ui.exFrom && d < ui.exFrom) return false;
        if (ui.exTo && d > ui.exTo) return false;
        if (ui.exOwner && t.owner !== ui.exOwner) return false;
        return true;
      })
      .sort(function (a, b) {
        return String(b.doneAt).localeCompare(String(a.doneAt));
      });

    var html = '';

    html +=
      '<div class="page-head"><div>' +
      '<div class="page-title">每日跟进记录导出</div>' +
      '<div class="page-desc">时间取「点击完成的时间」；同一天同一客户的多条事项会自动合并为一行并换行显示。</div>' +
      '</div><div class="page-actions">' +
      '<button class="btn btn-primary" data-action="export-xlsx">导出 XLSX</button>' +
      '</div></div>';

    html +=
      '<div class="card"><div class="card-head">' +
      '<div class="seg">' +
      '<button class="' + (ui.exMode === 'group' ? 'active' : '') + '" data-action="ex-mode" data-value="group">按客户合并</button>' +
      '<button class="' + (ui.exMode === 'detail' ? 'active' : '') + '" data-action="ex-mode" data-value="detail">明细逐条</button>' +
      '</div>' +
      '<button class="btn btn-sm" data-action="ex-today">今天</button>' +
      '<button class="btn btn-sm" data-action="ex-week">本周</button>' +
      '<button class="btn btn-sm" data-action="ex-month">本月</button>' +
      '<div class="card-head-actions">' +
      '<input class="input input-sm" type="date" data-action="ex-from" value="' + esc(ui.exFrom) + '">' +
      '<span style="color:var(--ink-4)">–</span>' +
      '<input class="input input-sm" type="date" data-action="ex-to" value="' + esc(ui.exTo) + '">' +
      '<select class="select select-sm" data-action="ex-owner"><option value="">全部负责人</option>' +
      ownerOptions(ui.exOwner) +
      '</select>' +
      '<button class="btn btn-sm" data-action="ex-reset">重置</button>' +
      '</div></div>';

    html += '<div class="card-body tight"><div class="table-scroll"><table class="tbl"><thead>';

    if (ui.exMode === 'group') {
      html +=
        '<tr><th style="width:120px">日期</th><th style="width:190px">客户</th><th>跟进事项</th>' +
        '<th style="width:74px">条数</th><th style="width:120px">负责人</th><th style="width:150px">备注</th></tr>';
    } else {
      html +=
        '<tr><th style="width:160px">完成时间</th><th style="width:190px">客户</th><th>事项</th>' +
        '<th style="width:120px">负责人</th><th style="width:150px">备注</th></tr>';
    }
    html += '</thead><tbody>';

    var rows = ui.exMode === 'group' ? groups : detail;
    if (!rows.length) {
      html +=
        '<tr><td colspan="' + (ui.exMode === 'group' ? 6 : 5) + '"><div class="empty"><div class="empty-mark">⤓</div>' +
        '<h3>暂无可导出的记录</h3><p>只有「已勾选完成」的任务才会进入导出结果，可调整上方日期范围。</p></div></td></tr>';
    } else if (ui.exMode === 'group') {
      groups.slice(0, 200).forEach(function (g) {
        html +=
          '<tr><td class="nowrap">' + esc(g.day) + '</td>' +
          '<td>' + esc(g.customer) + '</td>' +
          '<td class="task-content">' + esc(g.content) + '</td>' +
          '<td class="num">' + g.count + '</td>' +
          '<td>' + esc(g.owner || '—') + '</td>' +
          '<td>' + esc(g.note || '—') + '</td></tr>';
      });
    } else {
      detail.slice(0, 200).forEach(function (t) {
        html +=
          '<tr><td class="nowrap">' + esc(t.doneAt) + '</td>' +
          '<td>' + esc(t.customer || '—') + '</td>' +
          '<td class="task-content">' + esc(t.content) + '</td>' +
          '<td>' + esc(S.nameOfLogin(t.owner)) + '</td>' +
          '<td>' + esc(t.note || '—') + '</td></tr>';
      });
    }

    html += '</tbody></table></div>';
    html +=
      '<div class="pager"><span>' +
      (ui.exMode === 'group'
        ? '合并后 ' + groups.length + ' 行'
        : '共 ' + detail.length + ' 条记录') +
      (rows.length > 200 ? '（下方仅预览前 200 行，导出文件包含全部）' : '') +
      '</span></div></div></div>';

    // 备份与恢复
    var st = S.state;
    html +=
      '<div class="card"><div class="card-head"><div class="card-title">备份与恢复</div>' +
      '<div class="card-sub">数据保存在本机浏览器中，建议定期导出备份，或导出后发给同事做同步</div></div>' +
      '<div class="card-body">' +
      '<div class="stat-row" style="margin-bottom:16px">' +
      stat('客户', st.customers.length, '条记录', '') +
      stat('人员', st.personnel.length, '位', '') +
      stat('任务', st.tasks.length, '条（已完成 ' + st.tasks.filter(function (t) { return t.done; }).length + '）', 'is-ok') +
      stat('账号', st.accounts.length, '个', '') +
      '</div>' +
      '<div class="toolbar">' +
      '<button class="btn" data-action="backup-json">导出全量备份（JSON）</button>' +
      '<button class="btn" data-action="restore-json">从备份文件恢复</button>' +
      '<span class="grow"></span>' +
      '<button class="btn btn-danger" data-action="wipe-all">清空全部数据</button>' +
      '</div>' +
      '<div class="alert alert-warn" style="margin-top:14px;margin-bottom:0">' +
      '当前平台为纯前端版本，数据只保存在本机浏览器中。换电脑、换浏览器或清理浏览器数据都会丢失，请善用备份。<br>' +
      '若需要多人实时协作（跨人真实指派、云端同步），可随时改造成云端后端版本。' +
      '</div>' +
      '</div></div>';

    return html;
  }

  /* ============================ 界面5-1：客户 ============================ */

  function filteredCustomers() {
    var q = ui.search.trim().toLowerCase();
    return S.state.customers.filter(function (c) {
      if (ui.custSource && (c.source || '') !== ui.custSource) return false;
      if (ui.custLevel && (c.level || '') !== ui.custLevel) return false;
      if (ui.custSales && (c.sales || '') !== ui.custSales) return false;
      if (q) {
        var hay = [
          c.name, c.source, c.ctype, c.cmethod, c.sales, c.level, c.tags,
          c.wechat, c.decision, c.phone, c.address, c.brand, c.style,
          c.channel, c.pattern, c.trade, c.inquiry, c.lastContent,
        ]
          .join(' ')
          .toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }

  function viewCustomers() {
    var list = filteredCustomers();
    var pages = Math.max(1, Math.ceil(list.length / ui.custSize));
    if (ui.custPage > pages) ui.custPage = 1;
    var page = list.slice((ui.custPage - 1) * ui.custSize, ui.custPage * ui.custSize);

    var sources = uniqueSorted(S.state.customers.map(function (c) { return c.source; }));
    var levels = uniqueSorted(S.state.customers.map(function (c) { return c.level; }));
    var salesList = uniqueSorted(S.state.customers.map(function (c) { return c.sales; }));

    var html = '';

    html +=
      '<div class="page-head"><div>' +
      '<div class="page-title">数据源表格 · 客户</div>' +
      '<div class="page-desc">共 ' + S.state.customers.length + ' 位客户。支持导入 CRM 客户表格，重复的不会新建，只新增新的。</div>' +
      '</div><div class="page-actions">' +
      '<button class="btn" data-action="import-customers">导入 CRM 客户表格</button>' +
      '<button class="btn" data-action="export-customers">导出客户表</button>' +
      '<button class="btn btn-primary" data-action="add-customer">＋ 新增客户</button>' +
      '</div></div>';

    html +=
      '<div class="card"><div class="card-head">' +
      '<div class="toolbar grow">' +
      '<select class="select select-sm" data-action="cust-source"><option value="">全部来源</option>' +
      sources.map(function (s) {
        return '<option value="' + esc(s) + '"' + (s === ui.custSource ? ' selected' : '') + '>' + esc(s) + '</option>';
      }).join('') + '</select>' +
      '<select class="select select-sm" data-action="cust-level"><option value="">全部级别</option>' +
      levels.map(function (s) {
        return '<option value="' + esc(s) + '"' + (s === ui.custLevel ? ' selected' : '') + '>' + esc(s) + '</option>';
      }).join('') + '</select>' +
      '<select class="select select-sm" data-action="cust-sales"><option value="">全部业务员</option>' +
      salesList.map(function (s) {
        return '<option value="' + esc(s) + '"' + (s === ui.custSales ? ' selected' : '') + '>' + esc(s) + '</option>';
      }).join('') + '</select>' +
      '<button class="btn btn-sm" data-action="cust-reset">重置</button>' +
      '</div>' +
      '<div class="card-head-actions"><span class="card-sub">筛选结果 ' + list.length + ' 条</span></div>' +
      '</div>';

    html += '<div class="card-body tight"><div class="table-scroll"><table class="tbl"><thead><tr>';
    html +=
      '<th style="width:190px">客户名称</th><th style="width:110px">级别</th><th style="width:150px">来源</th>' +
      '<th style="width:150px">客户类型</th><th style="width:96px">业务员</th><th style="width:130px">微信</th>' +
      '<th style="width:110px">决策人</th><th style="width:120px">标签</th><th style="width:150px">最新跟进</th>' +
      '<th style="width:70px"></th>';
    html += '</tr></thead><tbody>';

    if (!page.length) {
      html +=
        '<tr><td colspan="10"><div class="empty" style="padding:36px"><h3>没有匹配的客户</h3><p>试试调整筛选条件或清空搜索关键词。</p></div></td></tr>';
    }

    page.forEach(function (c) {
      html +=
        '<tr data-id="' + esc(c.id) + '">' +
        '<td><b>' + esc(c.name) + '</b></td>' +
        '<td>' + levelTag(c.level) + '</td>' +
        '<td>' + esc(c.source || '—') + '</td>' +
        '<td>' + esc(c.ctype || '—') + '</td>' +
        '<td>' + esc(c.sales || '—') + '</td>' +
        '<td>' + esc(c.wechat || '—') + '</td>' +
        '<td>' + esc(c.decision || '—') + '</td>' +
        '<td>' + esc(c.tags || '—') + '</td>' +
        '<td class="nowrap">' + esc(c.lastFollow || '—') + '</td>' +
        '<td class="actions">' +
        '<button class="btn-icon" data-action="edit-customer" title="编辑">✎</button>' +
        '<button class="btn-icon" data-action="del-customer" title="删除">✕</button>' +
        '</td></tr>';
    });

    html += '</tbody></table></div>';

    if (list.length > ui.custSize) {
      html +=
        '<div class="pager"><span>第 ' + ui.custPage + ' / ' + pages + ' 页 · 每页 ' + ui.custSize + ' 条</span>' +
        '<span class="spacer"></span>' +
        '<button class="btn btn-sm" data-action="cust-prev"' + (ui.custPage <= 1 ? ' disabled' : '') + '>上一页</button>' +
        '<button class="btn btn-sm" data-action="cust-next"' + (ui.custPage >= pages ? ' disabled' : '') + '>下一页</button>' +
        '</div>';
    }

    html += '</div></div>';
    return html;
  }

  function levelTag(level) {
    if (!level) return '<span style="color:var(--ink-4)">—</span>';
    var short = String(level).replace('级', '');
    var cls = 'level-' + short.charAt(0).toUpperCase();
    return '<span class="tag">' + '<span class="level-dot ' + cls + '"></span>' + esc(level) + '</span>';
  }

  /* ============================ 界面5-2：人员架构表 ============================ */

  function viewPersonnel() {
    var people = S.state.personnel.filter(function (p) {
      return !ui.personRole || p.role === ui.personRole;
    });

    var html = '';
    html +=
      '<div class="page-head"><div>' +
      '<div class="page-title">数据源表格 · 人员架构表</div>' +
      '<div class="page-desc">按「主管理员 - 团队主管 - 销售 - 跟单」维护人员归属，可随时修改。</div>' +
      '</div><div class="page-actions">' +
      '<button class="btn" data-action="add-team">＋ 新增团队</button>' +
      '<button class="btn btn-primary" data-action="add-person">＋ 新增人员</button>' +
      '</div></div>';

    // 团队
    html +=
      '<div class="card"><div class="card-head"><div class="card-title">团队</div>' +
      '<div class="card-sub">共 ' + S.state.teams.length + ' 个</div></div><div class="card-body">' +
      '<div class="toolbar">';
    S.state.teams.forEach(function (t) {
      var n = S.state.personnel.filter(function (p) { return p.team === t.name; }).length;
      html +=
        '<span class="tag tag-brand">' + esc(t.name) + ' · ' + n + ' 人' +
        ' <button class="btn-icon" style="padding:0 2px;font-size:12px" data-action="del-team" data-id="' +
        esc(t.id) + '">✕</button></span>';
    });
    if (!S.state.teams.length) html += '<span style="color:var(--ink-4)">还没有团队</span>';
    html += '</div></div></div>';

    // 人员
    html +=
      '<div class="card"><div class="card-head">' +
      '<div class="seg">' +
      '<button class="' + (!ui.personRole ? 'active' : '') + '" data-action="person-role" data-value="">全部</button>' +
      Object.keys(S.ROLES).map(function (r) {
        return '<button class="' + (ui.personRole === r ? 'active' : '') + '" data-action="person-role" data-value="' +
          r + '">' + esc(roleLabel(r)) + '</button>';
      }).join('') +
      '</div>' +
      '<div class="card-head-actions"><span class="card-sub">' + people.length + ' 人</span></div>' +
      '</div>';

    html += '<div class="card-body tight"><div class="table-scroll"><table class="tbl"><thead><tr>';
    html +=
      '<th style="width:110px">姓名</th><th style="width:130px">登录账号</th><th style="width:110px">角色</th>' +
      '<th style="width:130px">团队</th><th style="width:110px">直属主管</th><th style="width:130px">工作手机</th>' +
      '<th style="width:120px">账号状态</th><th style="width:130px"></th>';
    html += '</tr></thead><tbody>';

    if (!people.length) {
      html += '<tr><td colspan="8"><div class="empty" style="padding:36px"><h3>还没有人员</h3><p>点击右上角「新增人员」开始搭建架构。</p></div></td></tr>';
    }

    people.forEach(function (p) {
      var acc = S.state.accounts.find(function (a) { return a.personnelId === p.id; }) ||
        S.state.accounts.find(function (a) { return a.login === p.login; });
      html +=
        '<tr data-id="' + esc(p.id) + '">' +
        '<td><b>' + esc(p.name) + '</b></td>' +
        '<td><code style="font-size:12.5px">' + esc(p.login) + '</code></td>' +
        '<td><span class="tag">' + esc(roleLabel(p.role)) + '</span></td>' +
        '<td>' + esc(p.team || '—') + '</td>' +
        '<td>' + esc(S.nameOfLogin(p.manager) || '—') + '</td>' +
        '<td>' + (p.phone ? esc(p.phone) : '<span style="color:var(--ink-4)">未填</span>') + '</td>' +
        '<td>' + (acc
          ? (acc.active ? '<span class="tag tag-ok">已开通</span>' : '<span class="tag tag-muted">已停用</span>')
          : '<span class="tag tag-warn">未开通</span>') + '</td>' +
        '<td class="actions">' +
        (acc ? '' : '<button class="btn btn-xs" data-action="make-account">开通账号</button>') +
        '<button class="btn-icon" data-action="edit-person" title="编辑">✎</button>' +
        '<button class="btn-icon" data-action="del-person" title="删除">✕</button>' +
        '</td></tr>';
    });

    html += '</tbody></table></div></div></div>';
    return html;
  }

  /* ============================ 界面5-3：账号登录 ============================ */

  function viewAccounts() {
    var accs = S.state.accounts;

    var html = '';
    html +=
      '<div class="page-head"><div>' +
      '<div class="page-title">数据源表格 · 账号登录</div>' +
      '<div class="page-desc">账号为姓名的拼音，密码为工作手机号。共 ' + accs.length + ' 个账号。</div>' +
      '</div><div class="page-actions">' +
      '<button class="btn btn-primary" data-action="add-account">＋ 新增账号</button>' +
      '</div></div>';

    html +=
      '<div class="alert alert-warn">登录校验在本机浏览器内完成，属于内部协作的便捷门禁，不等同于银行级安全防护。请勿将本页面截图外发。</div>';

    html += '<div class="card"><div class="card-body tight"><div class="table-scroll"><table class="tbl"><thead><tr>';
    html +=
      '<th style="width:110px">姓名</th><th style="width:140px">登录账号</th><th style="width:140px">密码（工作手机）</th>' +
      '<th style="width:110px">角色</th><th style="width:120px">团队</th><th style="width:100px">状态</th><th style="width:150px"></th>';
    html += '</tr></thead><tbody>';

    if (!accs.length) {
      html += '<tr><td colspan="7"><div class="empty" style="padding:36px"><h3>还没有账号</h3><p>点击右上角「新增账号」创建。</p></div></td></tr>';
    }

    accs.forEach(function (a) {
      html +=
        '<tr data-id="' + esc(a.id) + '">' +
        '<td><b>' + esc(a.name) + '</b></td>' +
        '<td><code style="font-size:12.5px">' + esc(a.login) + '</code></td>' +
        '<td class="nowrap">' + esc(a.phone || '—') + '</td>' +
        '<td><span class="tag">' + esc(roleLabel(a.role)) + '</span></td>' +
        '<td>' + esc(a.team || '—') + '</td>' +
        '<td>' + (a.active ? '<span class="tag tag-ok">启用</span>' : '<span class="tag tag-muted">停用</span>') + '</td>' +
        '<td class="actions">' +
        '<button class="btn btn-xs" data-action="edit-account">改手机号</button>' +
        '<button class="btn btn-xs" data-action="toggle-account">' + (a.active ? '停用' : '启用') + '</button>' +
        '<button class="btn-icon" data-action="del-account" title="删除">✕</button>' +
        '</td></tr>';
    });

    html += '</tbody></table></div></div></div>';
    return html;
  }

  /* ============================ 事件绑定 ============================ */

  function rerender() {
    render();
  }

  document.addEventListener('click', function (e) {
    var target = e.target && e.target.nodeType === 1 ? e.target : (e.target && e.target.parentElement);
    if (!target || !target.closest) return;

    var t = target.closest('[data-action]');
    var navChild = target.closest('[data-nav-child]');
    var navItem = target.closest('[data-nav]');

    if (navChild) {
      ui.view = navChild.getAttribute('data-nav-child');
      ui.search = '';
      rerender();
      return;
    }

    if (navItem && !navItem.classList.contains('dragging')) {
      var id = navItem.getAttribute('data-nav');
      var item = S.state.navOrder.find(function (n) {
        return n.id === id;
      });
      if (item && (item.children || []).length) {
        var willOpen = !ui.openNav[id];
        ui.openNav[id] = willOpen;
        if (willOpen) ui.view = (item.children[0] || {}).id || ui.view;
      } else {
        ui.view = id;
      }
      ui.search = '';
      rerender();
      return;
    }

    if (!t) {
      // 点击弹层外部关闭由 mask 自身处理
      return;
    }

    var action = t.getAttribute('data-action');
    var row = t.closest('tr');
    var rowId = row ? row.getAttribute('data-id') : null;

    switch (action) {
      /* ---- 通用 ---- */
      case 'close-modal':
        closeModal();
        break;
      case 'clear-search':
        ui.search = '';
        ui.todoPage = 1;
        ui.custPage = 1;
        rerender();
        break;
      case 'sign-out':
        S.signOut();
        ui.setupStage = 'login';
        rerender();
        break;
      case 'go-login':
        ui.setupStage = 'login';
        rerender();
        toast('账号创建成功，请登录', 'ok');
        break;

      /* ---- 今日抽屉 ---- */
      case 'open-today':
        ui.todayOpen = true;
        rerender();
        break;
      case 'close-today':
        // 点击抽屉内部不关闭，只有点到遮罩才关
        if (target.closest('.drawer')) break;
        ui.todayOpen = false;
        rerender();
        break;

      /* ---- 待办 ---- */
      case 'sort':
        ui.sortBy = t.getAttribute('data-value');
        ui.todoPage = 1;
        rerender();
        break;
      case 'toggle-showdone':
        ui.showDone = !ui.showDone;
        ui.todoPage = 1;
        rerender();
        break;
      case 'reset-filters':
        ui.fCustomer = '';
        ui.fFrom = '';
        ui.fTo = '';
        ui.fOwner = '';
        ui.todoPage = 1;
        rerender();
        break;
      case 'toggle-task':
        handleToggle(t, rowId);
        break;
      case 'del-task':
        if (rowId && confirm('确定删除这条任务？')) {
          S.removeTask(rowId);
          rerender();
          toast('已删除', 'ok');
        }
        break;
      case 'add-task':
        handleAddTask();
        break;
      case 'focus-add':
        var qc = el('quickContent');
        if (qc) {
          qc.focus();
          qc.scrollIntoView({ block: 'center', behavior: 'smooth' });
        }
        break;
      case 'todo-prev':
        ui.todoPage = Math.max(1, ui.todoPage - 1);
        rerender();
        break;
      case 'todo-next':
        ui.todoPage++;
        rerender();
        break;
      case 'export-todo':
        exportTodoList();
        break;

      /* ---- 导出 ---- */
      case 'ex-mode':
        ui.exMode = t.getAttribute('data-value');
        rerender();
        break;
      case 'ex-today':
        ui.exFrom = S.today();
        ui.exTo = S.today();
        rerender();
        break;
      case 'ex-week':
        ui.exFrom = weekStart();
        ui.exTo = S.today();
        rerender();
        break;
      case 'ex-month':
        ui.exFrom = monthStart();
        ui.exTo = S.today();
        rerender();
        break;
      case 'ex-reset':
        ui.exFrom = '';
        ui.exTo = '';
        ui.exOwner = '';
        rerender();
        break;
      case 'export-xlsx':
        exportXlsx();
        break;
      case 'backup-json':
        downloadText(
          S.exportAll(),
          '销售跟单平台备份_' + S.today() + '.json',
          'application/json;charset=utf-8'
        );
        toast('备份文件已生成', 'ok');
        break;
      case 'restore-json':
        pickFile('.json', function (file) {
          var reader = new FileReader();
          reader.onload = function () {
            var r = S.importAll(String(reader.result));
            if (r.error) return toast(r.error, 'err');
            rerender();
            toast('恢复完成', 'ok');
          };
          reader.readAsText(file, 'utf-8');
        });
        break;
      case 'wipe-all':
        if (
          confirm('将清空本机保存的客户、人员、任务与账号，且不可恢复。确定继续？') &&
          confirm('请再次确认：所有数据都会被删除。')
        ) {
          S.wipe();
          location.reload();
        }
        break;

      /* ---- 客户 ---- */
      case 'cust-prev':
        ui.custPage = Math.max(1, ui.custPage - 1);
        rerender();
        break;
      case 'cust-next':
        ui.custPage++;
        rerender();
        break;
      case 'cust-reset':
        ui.custSource = '';
        ui.custLevel = '';
        ui.custSales = '';
        ui.custPage = 1;
        rerender();
        break;
      case 'add-customer':
        openCustomerModal(null);
        break;
      case 'edit-customer':
        openCustomerModal(rowId);
        break;
      case 'del-customer':
        var cu = S.state.customers.find(function (x) { return x.id === rowId; });
        if (cu && confirm('确定删除客户「' + cu.name + '」？')) {
          S.removeCustomer(rowId);
          rerender();
          toast('已删除', 'ok');
        }
        break;
      case 'import-customers':
        importCustomersFlow();
        break;
      case 'export-customers':
        exportCustomers();
        break;

      /* ---- 人员 ---- */
      case 'add-person':
        openPersonModal(null);
        break;
      case 'edit-person':
        openPersonModal(rowId);
        break;
      case 'del-person':
        var pe = S.state.personnel.find(function (x) { return x.id === rowId; });
        if (pe && confirm('确定删除人员「' + pe.name + '」？其账号也会一并删除。')) {
          S.removePersonnel(rowId);
          rerender();
          toast('已删除', 'ok');
        }
        break;
      case 'add-team':
        var tn = prompt('新团队名称：');
        if (tn) {
          var r1 = S.addTeam(tn);
          if (r1.error) toast(r1.error, 'err');
          else {
            rerender();
            toast('团队已新增', 'ok');
          }
        }
        break;
      case 'del-team':
        var tid = t.getAttribute('data-id');
        var team = S.state.teams.find(function (x) { return x.id === tid; });
        if (team && confirm('删除团队「' + team.name + '」？成员不会被删除。')) {
          S.removeTeam(tid);
          rerender();
        }
        break;
      case 'person-role':
        ui.personRole = t.getAttribute('data-value');
        rerender();
        break;
      case 'make-account':
        var pd = S.state.personnel.find(function (x) { return x.id === rowId; });
        if (pd) openAccountModal(null, pd);
        break;

      /* ---- 账号 ---- */
      case 'add-account':
        openAccountModal(null, null);
        break;
      case 'edit-account':
        openAccountModal(rowId, null);
        break;
      case 'toggle-account':
        var ac = S.state.accounts.find(function (x) { return x.id === rowId; });
        if (ac) {
          S.updateAccount(rowId, { active: !ac.active });
          rerender();
          toast(ac.active ? '账号已停用' : '账号已启用', 'ok');
        }
        break;
      case 'del-account':
        var ac2 = S.state.accounts.find(function (x) { return x.id === rowId; });
        if (ac2 && confirm('确定删除账号「' + ac2.login + '」？')) {
          var r2 = S.removeAccount(rowId);
          if (r2.error) toast(r2.error, 'err');
          else {
            rerender();
            toast('已删除', 'ok');
          }
        }
        break;
    }
  });

  /* ---- 表单提交 ---- */
  document.addEventListener('submit', function (e) {
    var f = e.target;

    if (f.id === 'loginForm') {
      e.preventDefault();
      var login = el('loginName').value.trim();
      var pwd = el('loginPwd').value;
      var r = S.signIn(login, pwd);
      if (r.error) {
        el('loginErr').innerHTML = '<div class="alert alert-danger" style="margin-bottom:12px">' + esc(r.error) + '</div>';
        return;
      }
      ui.view = 'todo';
      toast('欢迎回来，' + r.data.name, 'ok');
      rerender();
      return;
    }

    if (f.id === 'setupForm') {
      e.preventDefault();
      var name = el('suName').value.trim();
      var suLogin = el('suLogin').value.trim().toLowerCase();
      var phone = el('suPhone').value.trim();
      var res = S.finishSetup({ name: name, login: suLogin, phone: phone, team: '', personnelId: '' });
      if (res.error) {
        ui.setupStage = 'error';
        ui.setupError = res.error;
        rerender();
        return;
      }
      ui.pendingSetup = { login: suLogin, phone: phone };
      ui.setupStage = 'pending';
      rerender();
      return;
    }

    if (f.id === 'assignForm') {
      e.preventDefault();
      var r3 = S.addTask({
        owner: el('asOwner').value,
        date: el('asDate').value,
        priority: el('asPriority').value,
        customer: el('asCustomer').value.trim(),
        note: el('asNote').value.trim(),
        content: el('asContent').value.trim(),
      });
      if (r3.error) return toast(r3.error, 'err');
      rerender();
      toast('任务已指派给 ' + S.nameOfLogin(r3.data.owner), 'ok');
      return;
    }

    if (f.id === 'customerForm') {
      e.preventDefault();
      var id = f.getAttribute('data-id');
      var payload = {
        name: el('cuName').value.trim(),
        level: el('cuLevel').value.trim(),
        source: el('cuSource').value.trim(),
        ctype: el('cuType').value.trim(),
        cmethod: el('cuMethod').value.trim(),
        sales: el('cuSales').value.trim(),
        wechat: el('cuWechat').value.trim(),
        decision: el('cuDecision').value.trim(),
        phone: el('cuPhone').value.trim(),
        address: el('cuAddress').value.trim(),
        brand: el('cuBrand').value.trim(),
        style: el('cuStyle').value.trim(),
        channel: el('cuChannel').value.trim(),
        pattern: el('cuPattern').value.trim(),
        trade: el('cuTrade').value.trim(),
        pricePref: el('cuPricePref').value.trim(),
        tags: el('cuTags').value.trim(),
        isNew: el('cuIsNew').value.trim(),
        production: el('cuProduction').value.trim(),
        design: el('cuDesign').value.trim(),
        salesPower: el('cuSalesPower').value.trim(),
        lastFollow: el('cuLastFollow').value.trim(),
        lastContent: el('cuLastContent').value.trim(),
        inquiry: el('cuInquiry').value.trim(),
      };
      var r4 = id ? S.updateCustomer(id, payload) : S.addCustomer(payload);
      if (r4.error) return toast(r4.error, 'err');
      closeModal();
      rerender();
      toast(id ? '客户已更新' : '客户已新增', 'ok');
      return;
    }

    if (f.id === 'personForm') {
      e.preventDefault();
      var pid = f.getAttribute('data-id');
      var pPayload = {
        name: el('peName').value.trim(),
        login: el('peLogin').value.trim().toLowerCase(),
        role: el('peRole').value,
        team: el('peTeam').value.trim(),
        manager: el('peManager').value,
        phone: el('pePhone').value.trim(),
        note: el('peNote').value.trim(),
      };
      var r5 = pid ? S.updatePersonnel(pid, pPayload) : S.addPersonnel(pPayload);
      if (r5.error) return toast(r5.error, 'err');
      // 若填了手机号且还没有账号，顺手开通
      var existing = S.state.accounts.find(function (a) {
        return a.personnelId === (r5.data && r5.data.id);
      });
      if (!existing && pPayload.phone && pPayload.login) {
        S.addAccount({
          personnelId: r5.data.id,
          name: pPayload.name,
          login: pPayload.login,
          phone: pPayload.phone,
          role: pPayload.role,
          team: pPayload.team,
        });
      }
      closeModal();
      rerender();
      toast(pid ? '人员已更新' : '人员已新增', 'ok');
      return;
    }

    if (f.id === 'accountForm') {
      e.preventDefault();
      var aid = f.getAttribute('data-id');
      var aPayload = {
        name: el('acName').value.trim(),
        login: el('acLogin').value.trim().toLowerCase(),
        phone: el('acPhone').value.trim(),
        role: el('acRole').value,
        team: el('acTeam').value.trim(),
      };
      var r6 = aid ? S.updateAccount(aid, aPayload) : S.addAccount(aPayload);
      if (r6.error) return toast(r6.error, 'err');
      closeModal();
      rerender();
      toast(aid ? '账号已更新' : '账号已创建', 'ok');
      return;
    }
  });

  /* ---- 输入 ---- */
  document.addEventListener('input', function (e) {
    var t = e.target;
    var action = t.getAttribute('data-action');

    if (t.id === 'globalSearch') {
      ui.search = t.value;
      ui.todoPage = 1;
      ui.custPage = 1;
      var clearBtn = document.querySelector('.search-box');
      if (clearBtn) clearBtn.classList.toggle('has-value', !!t.value);
      // 仅重绘主内容，保持输入焦点
      var main = el('main');
      if (main) main.innerHTML = renderMain(S.currentUser());
      return;
    }

    if (action === 'filter-from') {
      ui.fFrom = t.value;
      ui.todoPage = 1;
      rerender();
      return;
    }
    if (action === 'filter-to') {
      ui.fTo = t.value;
      ui.todoPage = 1;
      rerender();
      return;
    }
    if (action === 'ex-from') {
      ui.exFrom = t.value;
      rerender();
      return;
    }
    if (action === 'ex-to') {
      ui.exTo = t.value;
      rerender();
      return;
    }
  });

  document.addEventListener('change', function (e) {
    var t = e.target;
    var action = t.getAttribute('data-action');
    var row = t.closest('tr');

    if (action === 'filter-owner') {
      ui.fOwner = t.value;
      ui.todoPage = 1;
      rerender();
      return;
    }
    if (action === 'filter-customer') {
      ui.fCustomer = t.value;
      ui.todoPage = 1;
      rerender();
      return;
    }
    if (action === 'ex-owner') {
      ui.exOwner = t.value;
      rerender();
      return;
    }
    if (action === 'cust-source') {
      ui.custSource = t.value;
      ui.custPage = 1;
      rerender();
      return;
    }
    if (action === 'cust-level') {
      ui.custLevel = t.value;
      ui.custPage = 1;
      rerender();
      return;
    }
    if (action === 'cust-sales') {
      ui.custSales = t.value;
      ui.custPage = 1;
      rerender();
      return;
    }

    // 表格内联编辑
    var field = t.getAttribute('data-field');
    if (field && row) {
      var id = row.getAttribute('data-id');
      var patch = {};
      if (field === 'date') patch.date = S.normalizeDate(t.value);
      else patch[field] = t.value;
      S.updateTask(id, patch);
      if (field === 'owner') {
        rerender();
        toast('已改派给 ' + S.nameOfLogin(t.value), 'ok');
      }
      return;
    }
  });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') {
      if (el('modalMask')) closeModal();
      else if (ui.todayOpen) {
        ui.todayOpen = false;
        rerender();
      }
    }
    if (e.key === 'Enter' && e.target.id === 'quickContent') {
      e.preventDefault();
      handleAddTask();
    }
  });

  /* ---- 侧边栏拖拽排序 ---- */

  var dragFrom = null;

  document.addEventListener('dragstart', function (e) {
    var item = e.target.closest('.nav-item');
    if (!item) return;
    dragFrom = parseInt(item.getAttribute('data-index'), 10);
    item.classList.add('dragging');
    try {
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', String(dragFrom));
    } catch (err) {}
  });

  document.addEventListener('dragover', function (e) {
    var item = e.target.closest('.nav-item');
    if (!item || dragFrom == null) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    document.querySelectorAll('.nav-item.drop-target').forEach(function (n) {
      n.classList.remove('drop-target');
    });
    item.classList.add('drop-target');
  });

  document.addEventListener('dragend', function () {
    document.querySelectorAll('.nav-item').forEach(function (n) {
      n.classList.remove('dragging', 'drop-target');
    });
    dragFrom = null;
  });

  document.addEventListener('drop', function (e) {
    var item = e.target.closest('.nav-item');
    if (!item || dragFrom == null) return;
    e.preventDefault();
    var to = parseInt(item.getAttribute('data-index'), 10);
    if (isNaN(to) || to === dragFrom) return;
    var nav = S.state.navOrder.slice();
    var moved = nav.splice(dragFrom, 1)[0];
    nav.splice(to, 0, moved);
    S.state.navOrder = nav;
    S.save();
    dragFrom = null;
    rerender();
    toast('导航顺序已调整', 'ok');
  });

  /* ============================ 具体操作实现 ============================ */

  function handleAddTask() {
    var date = document.querySelector('[data-new="date"]');
    var customer = document.querySelector('[data-new="customer"]');
    var content = el('quickContent');
    var note = document.querySelector('[data-new="note"]');
    var owner = document.querySelector('[data-new="owner"]');
    if (!content || !content.value.trim()) return toast('请输入事项内容', 'warn');

    var r = S.addTask({
      date: date ? date.value : S.today(),
      customer: customer ? customer.value.trim() : '',
      content: content.value.trim(),
      note: note ? note.value.trim() : '',
      owner: owner ? owner.value : '',
    });
    if (r.error) return toast(r.error, 'err');
    rerender();
    var qc = el('quickContent');
    if (qc) qc.focus();
    toast('已新增', 'ok');
  }

  function handleToggle(inputEl, rowId) {
    var id = rowId || inputEl.getAttribute('data-id');
    if (!id) return;
    var wasDone = inputEl.checked;
    var r = S.toggleTask(id, wasDone);
    if (r.error) return toast(r.error, 'err');

    if (wasDone) {
      toast('「' + trim(r.data.content, 16) + '」已完成' + (ui.showDone ? '' : '并自动隐藏'), 'ok');
    } else {
      toast('已恢复为未完成', 'warn');
    }
    rerender();
  }

  function trim(s, n) {
    s = String(s || '');
    return s.length > n ? s.slice(0, n) + '…' : s;
  }

  function openCustomerModal(id) {
    var c = id
      ? S.state.customers.find(function (x) {
          return x.id === id;
        })
      : {};
    c = c || {};

    function section(title) {
      return (
        '<div class="drawer-section-title" style="margin-top:20px">' + esc(title) + '</div>'
      );
    }

    var body =
      '<form id="customerForm"' + (id ? ' data-id="' + esc(id) + '"' : '') + '>' +
      '<div class="grid-2">' +
      field('cuName', '客户名称 *', c.name) +
      field('cuLevel', '客户级别', c.level, '例如 A级 / D级') +
      '</div><div class="grid-2">' +
      field('cuSource', '客户来源', c.source) +
      field('cuType', '客户类型', c.ctype) +
      '</div><div class="grid-2">' +
      field('cuMethod', '客户对接方式', c.cmethod) +
      field('cuSales', '业务员', c.sales) +
      '</div>' +

      section('联系方式') +
      '<div class="grid-3">' +
      field('cuWechat', '微信', c.wechat) +
      field('cuDecision', '决策人', c.decision) +
      field('cuPhone', '电话', c.phone) +
      '</div>' +
      field('cuAddress', '地址', c.address) +

      section('经营画像') +
      '<div class="grid-3">' +
      field('cuBrand', '品牌', c.brand) +
      field('cuStyle', '产品风格', c.style) +
      field('cuChannel', '销售渠道', c.channel) +
      '</div><div class="grid-3">' +
      field('cuPattern', '业务模式', c.pattern) +
      field('cuTrade', '交易方式', c.trade) +
      field('cuPricePref', '价格偏好', c.pricePref) +
      '</div><div class="grid-3">' +
      field('cuTags', '客户标签', c.tags, '多个用逗号分隔') +
      field('cuIsNew', '当前年份新客', c.isNew, '是 / 否') +
      field('cuProduction', '生产能力', c.production) +
      '</div><div class="grid-3">' +
      field('cuDesign', '设计能力', c.design) +
      field('cuSalesPower', '销售能力', c.salesPower) +
      '<div></div>' +
      '</div>' +

      section('最近跟进') +
      '<div class="grid-2">' +
      field('cuLastFollow', '最新跟进日期', c.lastFollow, 'yyyy/mm/dd') +
      '<div></div>' +
      '</div>' +
      field('cuLastContent', '最新跟进内容', c.lastContent) +
      field('cuInquiry', '询盘内容', c.inquiry) +
      '<div style="display:none"><button type="submit"></button></div>' +
      '</form>';
    var foot =
      '<button class="btn" data-action="close-modal">取消</button>' +
      '<button class="btn btn-primary" onclick="document.getElementById(\'customerForm\').requestSubmit()">保存</button>';
    openModal(id ? '编辑客户' : '新增客户', body, foot, true);
  }

  function openPersonModal(id) {
    var p = id
      ? S.state.personnel.find(function (x) {
          return x.id === id;
        })
      : {};
    p = p || {};

    var roleOpts = Object.keys(S.ROLES)
      .map(function (r) {
        return '<option value="' + r + '"' + (p.role === r ? ' selected' : '') + '>' + esc(roleLabel(r)) + '</option>';
      })
      .join('');

    var teamOpts =
      '<option value="">未分组</option>' +
      S.state.teams
        .map(function (t) {
          return '<option value="' + esc(t.name) + '"' + (p.team === t.name ? ' selected' : '') + '>' + esc(t.name) + '</option>';
        })
        .join('');

    var managerOpts =
      '<option value="">未指定</option>' +
      S.state.personnel
        .filter(function (x) {
          return x.role === 'manager' || x.role === 'admin';
        })
        .map(function (x) {
          return '<option value="' + esc(x.login) + '"' + (p.manager === x.login ? ' selected' : '') + '>' +
            esc(x.name) + '（' + esc(roleLabel(x.role)) + '）</option>';
        })
        .join('');

    var body =
      '<form id="personForm"' + (id ? ' data-id="' + esc(id) + '"' : '') + '>' +
      '<div class="grid-2">' +
      field('peName', '姓名 *', p.name) +
      field('peLogin', '登录账号 *', p.login, '名字拼音，如 zhangwei') +
      '</div><div class="grid-2">' +
      '<div class="field"><label>角色</label><select class="select" id="peRole">' + roleOpts + '</select></div>' +
      '<div class="field"><label>团队</label><select class="select" id="peTeam">' + teamOpts + '</select></div>' +
      '</div><div class="grid-2">' +
      '<div class="field"><label>直属主管</label><select class="select" id="peManager">' + managerOpts + '</select></div>' +
      field('pePhone', '工作手机', p.phone, '填写后将自动开通账号') +
      '</div>' +
      field('peNote', '备注', p.note) +
      '<div style="display:none"><button type="submit"></button></div>' +
      '</form>';
    var foot =
      '<button class="btn" data-action="close-modal">取消</button>' +
      '<button class="btn btn-primary" onclick="document.getElementById(\'personForm\').requestSubmit()">保存</button>';
    openModal(id ? '编辑人员' : '新增人员', body, foot);
  }

  function openAccountModal(id, presetPerson) {
    var a = id
      ? S.state.accounts.find(function (x) {
          return x.id === id;
        })
      : {};
    a = a || {};
    var pre = presetPerson || {};

    var roleOpts = Object.keys(S.ROLES)
      .map(function (r) {
        var sel = (a.role || pre.role) === r ? ' selected' : '';
        return '<option value="' + r + '"' + sel + '>' + esc(roleLabel(r)) + '</option>';
      })
      .join('');

    var body =
      '<form id="accountForm"' + (id ? ' data-id="' + esc(id) + '"' : '') + '>' +
      '<div class="alert alert-info">账号规则（来自 SOP）：<b>账号为姓名的拼音，密码为工作手机号</b>。</div>' +
      '<div class="grid-2">' +
      field('acName', '姓名 *', a.name || pre.name) +
      field('acLogin', '登录账号 *', a.login || pre.login, '拼音') +
      '</div><div class="grid-2">' +
      field('acPhone', '工作手机（即密码）*', a.phone || pre.phone) +
      '<div class="field"><label>角色</label><select class="select" id="acRole">' + roleOpts + '</select></div>' +
      '</div>' +
      field('acTeam', '团队', a.team || pre.team) +
      '<div style="display:none"><button type="submit"></button></div>' +
      '</form>';
    var foot =
      '<button class="btn" data-action="close-modal">取消</button>' +
      '<button class="btn btn-primary" onclick="document.getElementById(\'accountForm\').requestSubmit()">保存</button>';
    openModal(id ? '编辑账号' : '新增账号', body, foot);
  }

  function field(id, label, value, hint) {
    return (
      '<div class="field"><label>' + esc(label) +
      (hint ? '<span class="hint">' + esc(hint) + '</span>' : '') +
      '</label><input class="input" id="' + id + '" value="' + esc(value == null ? '' : value) + '"></div>'
    );
  }

  function pickFile(accept, cb) {
    var inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = accept;
    inp.onchange = function () {
      if (inp.files && inp.files[0]) cb(inp.files[0]);
    };
    inp.click();
  }

  /* ---- 客户导入 ---- */

  function importCustomersFlow() {
    pickFile('.xlsx,.xls,.csv', function (file) {
      if (/\.xls$/i.test(file.name)) {
        return toast('旧版 .xls 暂不支持，请在 Excel 中另存为 .xlsx 或 .csv 后再导入', 'warn');
      }
      toast('正在解析 ' + file.name + '…');
      READ.readFile(file)
        .then(function (res) {
          var rows = res.rows.filter(function (r) {
            return r.some(function (c) {
              return String(c || '').trim();
            });
          });
          if (!rows.length) return toast('文件里没有可读取的内容', 'warn');

          var mapped = S.mapCustomerRows(rows);
          if (mapped.error) return toast(mapped.error + '，请确认是 CRM 导出的客户表格', 'err');
          if (!mapped.payload.length) return toast('没有解析到有效的客户记录', 'warn');

          var r = S.importCustomers(mapped.payload);
          rerender();
          openModal(
            '导入完成',
            '<div class="stat-row">' +
              stat('新增', r.data.added, '条新客户', 'is-ok') +
              stat('跳过', r.data.skipped, '条已存在', 'is-warn') +
              stat('客户总数', r.data.total, '条', '') +
            '</div>' +
            '<div class="alert alert-info">已按客户名称去重：重复的不会新建，只新增表格里新出现的客户。</div>' +
            '<div class="card-sub">来源文件：' + esc(res.sheetName) + ' · 读取 ' + mapped.rows + ' 行' +
            (mapped.skippedNoName ? ' · ' + mapped.skippedNoName + ' 行因缺少名称被忽略' : '') +
            '</div>',
            '<button class="btn btn-primary" data-action="close-modal">知道了</button>'
          );
        })
        .catch(function (err) {
          console.error(err);
          toast('解析失败：' + (err && err.message ? err.message : '未知错误'), 'err');
        });
    });
  }

  /* ---- 导出实现 ---- */

  function exportTodoList() {
    var list = filteredTasks();
    var rows = list.map(function (t) {
      return [
        S.dateOnly(t.date),
        t.customer,
        t.content,
        t.done ? '已完成' : '未完成',
        t.note,
        S.nameOfLogin(t.owner),
        t.creator ? S.nameOfLogin(t.creator) : '',
      ];
    });
    var bytes = XLSX.build({
      sheetName: '日常待办',
      freeze: true,
      columns: [
        { header: '日期', width: 13 },
        { header: '客户', width: 22 },
        { header: '事项', width: 46 },
        { header: '是否完成', width: 11 },
        { header: '备注', width: 22 },
        { header: '负责人', width: 12 },
        { header: '创建人', width: 12 },
      ],
      rows: rows,
    });
    XLSX.download(bytes, '日常待办_' + S.today() + '.xlsx');
    toast('已导出 ' + rows.length + ' 条', 'ok');
  }

  function exportXlsx() {
    if (ui.exMode === 'group') {
      var groups = S.groupDoneTasks({ from: ui.exFrom, to: ui.exTo, owner: ui.exOwner });
      if (!groups.length) return toast('当前范围内没有已完成的记录', 'warn');
      var rows = groups.map(function (g) {
        return [
          g.day,
          g.customer,
          g.content,
          g.count,
          g.owner,
          g.note,
        ];
      });
      var bytes = XLSX.build({
        sheetName: '每日跟进记录',
        freeze: true,
        columns: [
          { header: '日期', width: 13 },
          { header: '客户', width: 24 },
          { header: '跟进事项', width: 52 },
          { header: '条数', width: 8 },
          { header: '负责人', width: 14 },
          { header: '备注', width: 24 },
        ],
        rows: rows.map(function (r) {
          return [r[0], r[1], r[2], r[3], r[4], r[5]];
        }),
      });
      var name =
        '每日跟进记录_' +
        (ui.exFrom || '全部') +
        (ui.exTo && ui.exTo !== ui.exFrom ? '至' + ui.exTo : '') +
        '.xlsx';
      XLSX.download(bytes, name);
      toast('已导出 ' + groups.length + ' 行（同客户已合并）', 'ok');
      return;
    }

    // 明细
    var detail = S.state.tasks
      .filter(function (t) {
        if (!t.done || !t.doneAt) return false;
        var d = S.dateOnly(t.doneAt);
        if (ui.exFrom && d < ui.exFrom) return false;
        if (ui.exTo && d > ui.exTo) return false;
        if (ui.exOwner && t.owner !== ui.exOwner) return false;
        return true;
      })
      .sort(function (a, b) {
        return String(a.doneAt).localeCompare(String(b.doneAt));
      });

    if (!detail.length) return toast('当前范围内没有已完成的记录', 'warn');

    var bytes2 = XLSX.build({
      sheetName: '跟进明细',
      freeze: true,
      columns: [
        { header: '完成时间', width: 19 },
        { header: '日期', width: 13 },
        { header: '客户', width: 24 },
        { header: '事项', width: 46 },
        { header: '负责人', width: 12 },
        { header: '创建人', width: 12 },
        { header: '备注', width: 22 },
      ],
      rows: detail.map(function (t) {
        return [
          t.doneAt,
          S.dateOnly(t.date),
          t.customer,
          t.content,
          S.nameOfLogin(t.owner),
          S.nameOfLogin(t.creator),
          t.note,
        ];
      }),
    });
    XLSX.download(bytes2, '跟进明细_' + S.today() + '.xlsx');
    toast('已导出 ' + detail.length + ' 条明细', 'ok');
  }

  function exportCustomers() {
    var list = filteredCustomers();
    if (!list.length) return toast('没有可导出的客户', 'warn');
    var bytes = XLSX.build({
      sheetName: '客户',
      freeze: true,
      columns: [
        { header: '客户名称', width: 24 },
        { header: '客户级别', width: 11 },
        { header: '客户来源', width: 20 },
        { header: '客户类型', width: 20 },
        { header: '客户对接方式', width: 16 },
        { header: '业务员', width: 12 },
        { header: '微信', width: 20 },
        { header: '决策人', width: 16 },
        { header: '电话', width: 16 },
        { header: '地址', width: 26 },
        { header: '客户标签', width: 16 },
        { header: '品牌', width: 16 },
        { header: '产品风格', width: 14 },
        { header: '销售渠道', width: 14 },
        { header: '业务模式', width: 14 },
        { header: '交易方式', width: 12 },
        { header: '价格偏好', width: 14 },
        { header: '当前年份新客', width: 13 },
        { header: '生产能力', width: 12 },
        { header: '设计能力', width: 12 },
        { header: '销售能力', width: 12 },
        { header: '最新跟进日期', width: 19 },
        { header: '最新跟进内容', width: 36 },
        { header: '询盘内容', width: 30 },
      ],
      rows: list.map(function (c) {
        return [
          c.name,
          c.level,
          c.source,
          c.ctype,
          c.cmethod,
          c.sales,
          c.wechat,
          c.decision,
          c.phone,
          c.address,
          c.tags,
          c.brand,
          c.style,
          c.channel,
          c.pattern,
          c.trade,
          c.pricePref,
          c.isNew,
          c.production,
          c.design,
          c.salesPower,
          c.lastFollow,
          c.lastContent,
          c.inquiry,
        ];
      }),
    });
    XLSX.download(bytes, '客户表_' + S.today() + '.xlsx');
    toast('已导出 ' + list.length + ' 位客户', 'ok');
  }

  /* ============================ 启动 ============================ */

  function boot() {
    S.load();
    var seeded = S.seedIfEmpty();
    if (seeded) S.save();
    if (!S.state.meta.initialized) ui.setupStage = 'login';
    render();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
