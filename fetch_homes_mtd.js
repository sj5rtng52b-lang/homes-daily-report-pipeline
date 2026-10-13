/* Homes店长大盘取数脚本 —— 在 quark.alibaba.com 报表页(reportId=312793)的 DevTools/浏览器自动化上下文执行
 * 使用前替换占位符: 20260914=统计日期YYYYMMDD  8790=本地接收端口  mtd=周期(mtd月度累计 / 1d单日)
 * 作用: 拉取4个widget数据(品类-部门/商品管理-部门/整体表现/分渠道成交)，POST到本地 homes_server.py
 * 返回: 'POSTED ...' 成功；'ERR ...' 失败
 *
 * 【2026-09-11 改造】传输层由 fetch 改为 XMLHttpRequest。
 * 原因: 实测报表页上下文 fetch() 偶发全面失效(同域 GET 也报 Failed to fetch，
 *       而 XHR/顶层导航正常，疑似扩展或页面脚本钩住 window.fetch)，XHR 不受影响。
 * 本地回环 POST 使用同步 XHR，避免服务器收数即退出导致的连接竞态误报。
 */
(function () {
  var DATE = '20260914', PORT = '8790', PERIOD = 'mtd';
  var m = location.href.match(/QUARK_PARAMS=([^&]*)/);
  if (!m) return 'ERR_NO_QUARK_PARAMS ' + location.href.slice(0, 120);
  var QP = m[1];
  function strip(t) { var s = t.indexOf('('), e = t.lastIndexOf(')'); return JSON.parse(t.slice(s + 1, e)); }
  function xhrPostAsync(url, body) {
    return new Promise(function (resolve, reject) {
      var x = new XMLHttpRequest();
      x.open('POST', url, true);
      x.withCredentials = true;
      x.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      x.onload = function () { resolve(x.responseText); };
      x.onerror = function () { reject(new Error('xhr network fail ' + url.slice(0, 80))); };
      x.timeout = 60000;
      x.ontimeout = function () { reject(new Error('xhr timeout ' + url.slice(0, 80))); };
      x.send(body);
    });
  }
  function fetchW(wid, enc, gn, conds, cb) {
    var body = 'param=' + encodeURIComponent(JSON.stringify({ conditionList: conds, groupNo: gn })) +
      '&groupNo=' + gn + '&widgetId=' + wid + '&encryptWidgetId=' + enc +
      '&externalFlag=true&reportId=312793&env=pro&requestType=view&dtsJobId=&QUARK_PARAMS=' + QP;
    return xhrPostAsync('https://quark.alibaba.com/.dox?action=WidgetAction&event_submit_doGetWidgetData=true&callback=' + cb + '&_ts_=auto_' + wid, body)
      .then(function (t) {
        var j = strip(t);
        if (j.returnCode !== 0) throw new Error('widget ' + wid + ' rc=' + j.returnCode + ' ' + (j.returnMessage || ''));
        return j.returnValue.datas.datas;
      });
  }
  var cat = [
    { conditionId: '494366', v1: DATE, v2: DATE }, { conditionId: '494367', v1: DATE, v2: DATE },
    { conditionId: '473800', v1: [PERIOD] }, { conditionId: '473513', v1: [] },
    { conditionId: '473514', v1: ['19005'] }, { conditionId: '473515', v1: [] },
    { conditionId: '473597', v1: [] }, { conditionId: '473598', v1: [] }, { conditionId: '473516', v1: '' }
  ];
  var st = [
    { conditionId: '473566', v1: [PERIOD] },
    { conditionId: '494364', v1: DATE, v2: DATE }, { conditionId: '494365', v1: DATE, v2: DATE },
    { conditionId: '473509', v1: [] }, { conditionId: '473510', v1: ['19005'] }, { conditionId: '473511', v1: '' }
  ];
  return Promise.all([
    fetchW('497734', '75a35406af74613c', '14336961', cat, 'ca'),
    fetchW('497737', 'bff564290024995a', '14336961', cat, 'cb'),
    fetchW('497731', '771f514d9159da27', '14336959', st, 'cc'),
    fetchW('497780', 'bd70adf8fc9f05b3', '14336959', st, 'cd')
  ]).then(function (rs) {
    var payload = JSON.stringify({ date: DATE, period: PERIOD, cat: rs[0], loss: rs[1], overall: rs[2], channel: rs[3] });
    // 同步 XHR 提交本地回环，规避服务器收数即退出的连接竞态误报
    var x = new XMLHttpRequest();
    try {
      x.open('POST', 'http://127.0.0.1:' + PORT + '/payload', false);
      x.send(payload);
      if (x.status === 200) return 'POSTED ' + payload.length + 'B -> ' + x.responseText;
      return 'ERR_LOCAL_HTTP_' + x.status;
    } catch (e) {
      if (x.responseText && x.responseText.indexOf('RECEIVED') >= 0) {
        return 'POSTED ' + payload.length + 'B -> ' + x.responseText;
      }
      return 'ERR_LOCAL_POST ' + (e && e.message ? e.message : e);
    }
  }).catch(function (e) { return 'ERR ' + e.message; });
})()
