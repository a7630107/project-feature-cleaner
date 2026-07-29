(function(){
  "use strict";
  var DEFAULT_TPL = [
    {cat:"电气管线~电缆", fmt:"1、名称:{材质}\n2、规格型号:{规格型号}\n3、敷设方式:管内或桥架内"},
    {cat:"电气管线~电线", fmt:"1、名称:管内穿线\n2、型号、规格:{规格型号}"},
    {cat:"电气管线~配管", fmt:"1、名称:{材质}\n2、规格型号:{规格型号}\n3、配置形式:砖、混凝土结构暗配"},
    {cat:"电气管线~桥架", fmt:"1、名称:{材质}\n2、规格型号:{规格型号}\n3、接地方式:-25*4热镀锌扁钢跨接接地"}
  ];
  var RAW=null, CLEANED=[], BILL=[], TEMPLATES=DEFAULT_TPL.map(function(x){ return {cat:x.cat, fmt:x.fmt, qtyNameMap:"", qtyMap:""}; });

  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s==null?"":s).replace(/[&<>]/g, function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;"}[c]; }); }
  function escAttr(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
  function colLetter(n){ var s=""; while(n>0){ var m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; }

  function showSec(id){
    ["sec-import","sec-clean","sec-bill","sec-tpl"].forEach(function(s){
      $(s).style.display = (s===id) ? "block" : "none";
    });
    Array.prototype.forEach.call(document.querySelectorAll(".nav-item"), function(el){
      el.classList.toggle("active", el.getAttribute("data-sec")===id);
    });
  }

  function initNav(){
    Array.prototype.forEach.call(document.querySelectorAll(".nav-item"), function(el){
      el.addEventListener("click", function(){ showSec(el.getAttribute("data-sec")); });
    });
    $("sidebar-toggle").addEventListener("click", function(){
      $("sidebar").classList.toggle("collapsed");
    });
  }

  function handleSummaryFile(file){
    var reader = new FileReader();
    reader.onload = function(e){
      try{
        var data = new Uint8Array(e.target.result);
        RAW = XLSX.read(data, {type:"array"});
        var list = RAW.SheetNames.map(function(n){
          var ws = RAW.Sheets[n];
          var cnt = (ws && ws["!ref"]) ? (XLSX.utils.decode_range(ws["!ref"]).e.r) : 0;
          return n + "（" + cnt + " 行）";
        });
        $("import-status").innerHTML = "已读取工作簿，共 " + RAW.SheetNames.length + " 张表：<br>" + list.join("<br>");
        $("btn-clean").disabled = false;
      }catch(err){ alert("读取失败：" + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleTplFile(file){
    var reader = new FileReader();
    reader.onload = function(ev){
      try{
        var wb = XLSX.read(new Uint8Array(ev.target.result), {type:"array"});
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
        var hdr = aoa[0] || [];
        var iCat = -1, iFmt = -1, iQtyName = -1, iQtyMap = -1;
        for(var c=0;c<hdr.length;c++){
          var h = String(hdr[c]||"");
          if(h.indexOf("分类")>=0) iCat = c;
          if(h.indexOf("项目特征格式")>=0) iFmt = c;
          if(h.indexOf("工程量名称")>=0 || h.toLowerCase().indexOf("qtynamemap")>=0) iQtyName = c;
          if(h.indexOf("工程量映射")>=0 || h.toLowerCase().indexOf("qtymap")>=0) iQtyMap = c;
        }
        if(iCat<0) iCat = 0;
        if(iFmt<0) iFmt = (hdr.length>=4 ? 3 : 1); // 兼容旧列：分类/参考/格式
        var t = [];
        for(var i=1;i<aoa.length;i++){
          var row = aoa[i];
          var cat = String(row[iCat]||"").trim();
          var fmt = String(row[iFmt]||"");
          if(!cat) continue;
          t.push({cat:cat, fmt:fmt,
            qtyNameMap:(iQtyName>=0 ? String(row[iQtyName]||"") : ""),
            qtyMap:(iQtyMap>=0 ? String(row[iQtyMap]||"") : "")});
        }
        if(t.length){ TEMPLATES = t; renderTpl(); onTemplatesChanged(); $("tpl-status").textContent = "已导入模板 " + t.length + " 条，已覆盖当前模板库。"; }
        else { $("tpl-status").textContent = "未从文件中解析到模板（请确认第一张表含 分类 / 项目特征格式 两列）。"; }
      }catch(err){ alert("模板导入失败：" + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }

  function renderTpl(){
    var html = "<table><thead><tr><th>分类（构建类型~计算项目）</th><th>项目特征格式（用 {列名} 引用表头；{对应工程量名称表头名} 显示工程量名称，{对应工程量表头名} 显示工程量数值）</th><th>工程量名称映射<br><span class='tpl-hint'>默认 工程量名称 列；填 {表头名} 如 {工程量名称}</span></th><th>工程量映射<br><span class='tpl-hint'>默认 工程量 列；填 {表头名} 如 {长度合计}</span></th><th>操作</th></tr></thead><tbody>";
    if(!TEMPLATES.length){
      html += "<tr><td colspan='5' style='text-align:center;color:#888'>暂无模板，点「新增模板」添加一行。</td></tr>";
    }
    TEMPLATES.forEach(function(t, i){
      html += "<tr data-i='" + i + "'>";
      html += "<td><input class='tpl-cat' data-i='" + i + "' value='" + escAttr(t.cat) + "'></td>";
      html += "<td><textarea class='tpl-fmt' data-i='" + i + "'>" + esc(t.fmt) + "</textarea></td>";
      html += "<td><input class='tpl-qtyname' data-i='" + i + "' value='" + escAttr(t.qtyNameMap||"") + "' placeholder='{工程量名称}'></td>";
      html += "<td><input class='tpl-qty' data-i='" + i + "' value='" + escAttr(t.qtyMap||"") + "' placeholder='{长度合计}'></td>";
      html += "<td><button class='tpl-del' data-i='" + i + "'>删除</button></td>";
      html += "</tr>";
    });
    html += "</tbody></table>";
    var box = $("tpl-table");
    if(box) box.innerHTML = html;
  }

  function exportTpl(){
    if(!TEMPLATES.length){ alert("暂无模板可导出"); return; }
    var aoa = [["分类","项目特征格式","工程量名称","工程量映射"]];
    TEMPLATES.forEach(function(t){ aoa.push([t.cat, t.fmt, t.qtyNameMap||"", t.qtyMap||""]); });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "模板库");
    XLSX.writeFile(wb, "项目特征模板库.xlsx");
  }

  // 模板被修改/新增/删除/导入/重置后：若清单已生成，则自动按最新模板重算并刷新清单表
  function onTemplatesChanged(){
    if(!BILL.length) return;            // 还没生成清单，无需更新
    BILL = mergeBill(genBill());        // 依据当前 TEMPLATES + CLEANED 重算，并按项目特征合并汇总
    renderBill();                       // 重绘清单表（若当前在清单表，立即看到更新）
    var st = $("bill-status");
    if(st) st.textContent = "模板已修改，清单已自动更新，当前 " + BILL.length + " 行。";
  }

  // split at most maxsplits times, preserving the remainder (matches Python split(sep, maxsplit))
  function splitMax(s, sep, maxsplits){
    s = String(s);
    if(maxsplits <= 0) return [s];
    var arr = [], idx = 0, count = 0, pos;
    while(count < maxsplits && (pos = s.indexOf(sep, idx)) >= 0){
      arr.push(s.slice(idx, pos));
      idx = pos + sep.length;
      count++;
    }
    arr.push(s.slice(idx));
    return arr;
  }

  function sheetToMatrix(ws){
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var R = range.e.r - range.s.r + 1;
    var C = range.e.c - range.s.c + 1;
    var grid = [];
    var r, c;
    for(r=0;r<R;r++){ grid.push(new Array(C).fill(null)); }
    for(r=0;r<R;r++) for(c=0;c<C;c++){
      var addr = XLSX.utils.encode_cell({r:r+range.s.r, c:c+range.s.c});
      var cell = ws[addr];
      grid[r][c] = cell ? (cell.v===undefined ? null : cell.v) : null;
    }
    var merges = ws["!merges"] || [];
    merges.forEach(function(m){
      var tv = grid[m.s.r-range.s.r][m.s.c-range.s.c];
      for(var rr=m.s.r; rr<=m.e.r; rr++) for(var cc=m.s.c; cc<=m.e.c; cc++){
        grid[rr-range.s.r][cc-range.s.c] = tv;
      }
    });
    return grid;
  }

  function cleanWorkbook(wb){
    var out = [];
    wb.SheetNames.forEach(function(name){
      var ws = wb.Sheets[name];
      if(!ws || !ws["!ref"]) return;
      var grid = sheetToMatrix(ws);
      var headers = grid[0].map(function(h){ return h==null ? "" : String(h); });
      var data = grid.slice(1);
      var bt = name.replace(/汇总表/g, "").replace(/工程量/g, "");
      headers.unshift("构建类型");
      data.forEach(function(row){ row.unshift(bt); });
      // split dash headers
      var nh = [], plan = [];
      headers.forEach(function(h){
        if(h && h.indexOf("-")>=0){
          var parts = h.split("-");
          parts.forEach(function(p){ nh.push(p); });
          plan.push(parts.length);
        } else { nh.push(h); plan.push(0); }
      });
      var nd = [];
      data.forEach(function(row){
        var nr = [];
        for(var j=0;j<headers.length;j++){
          var n = plan[j];
          var v = row[j];
          if(n>0){
            if(v==null){ for(var k=0;k<n;k++) nr.push(""); }
            else { var sp = splitMax(v, "-", n-1); while(sp.length<n) sp.push(""); sp.forEach(function(x){ nr.push(x); }); }
          } else { nr.push(v==null ? "" : v); }
        }
        nd.push(nr);
      });
      out.push({name:name, headers:nh, rows:nd});
    });
    return out;
  }

  function renderClean(){
    var sel = $("clean-sheet");
    var idx = parseInt(sel.value, 10);
    if(isNaN(idx) || !CLEANED[idx]) return;
    var sheet = CLEANED[idx];
    var q = $("clean-search").value.trim().toLowerCase();
    var html = "<table><thead><tr>";
    sheet.headers.forEach(function(h){ html += "<th>" + esc(h) + "</th>"; });
    html += "</tr></thead><tbody>";
    sheet.rows.forEach(function(row, i){
      var hay = row.map(function(x){ return x==null ? "" : String(x); }).join(" ").toLowerCase();
      if(q && hay.indexOf(q) < 0) return;
      html += "<tr data-row='" + (i+2) + "'>";
      row.forEach(function(cell){ html += "<td>" + esc(cell) + "</td>"; });
      html += "</tr>";
    });
    html += "</tbody></table>";
    $("clean-table").innerHTML = html;
  }

  function fillTemplate(fmt, row, hi, iQtyName, iQty){
    return fmt.replace(/\{([^}]*)\}/g, function(m, tok){
      tok = tok.trim();
      // 特殊占位符 {对应工程量名称表头名} → 当前行工程量名称列的值（受模板「工程量名称」映射控制）
      if(tok === "对应工程量名称表头名" && iQtyName != null && row[iQtyName] != null && String(row[iQtyName]).trim() !== ""){
        return String(row[iQtyName]);
      }
      // 特殊占位符 {对应工程量表头名} → 当前行工程量列的值（受模板「工程量映射」控制）
      if(tok === "对应工程量表头名" && iQty != null && row[iQty] != null && String(row[iQty]).trim() !== ""){
        return String(row[iQty]);
      }
      if(hi[tok]!=null && row[hi[tok]]!=null && String(row[hi[tok]]).trim() !== ""){
        return String(row[hi[tok]]);
      }
      return "{" + tok + "}";
    });
  }

  function genBill(){
    var bill = [];
    TEMPLATES.forEach(function(t){
      var parts = t.cat.split("~");
      var tbt = parts[0].trim();
      var tcp = parts.length > 1 ? parts[1].trim() : null;
      // 工程量名称映射：模板的 qtyNameMap 字段（支持 {HeaderName} 语法）→ 默认 "工程量名称"
      var qnm = (t.qtyNameMap == null ? "" : String(t.qtyNameMap).trim());
      var qnmMatch = qnm.match(/^\{(.+)\}$/);
      var qtyNameColName = (qnmMatch ? qnmMatch[1].trim() : (qnm || "工程量名称"));
      // 工程量映射：模板的 qtyMap 字段（支持 {HeaderName} 语法）→ 默认 "工程量"
      var qm = (t.qtyMap == null ? "" : String(t.qtyMap).trim());
      var qmMatch = qm.match(/^\{(.+)\}$/);
      var qtyColName = (qmMatch ? qmMatch[1].trim() : (qm || "工程量"));
      CLEANED.forEach(function(sheet){
        var hi = {}; sheet.headers.forEach(function(h, i){ hi[h] = i; });
        var iBT = hi["构建类型"], iCP = hi["计算项目"], iUnit = hi["单位"];
        var iQtyName = hi[qtyNameColName];  // 找不到则 undefined
        var iQty = hi[qtyColName];          // 找不到则 undefined
        if(iBT == null) return;
        sheet.rows.forEach(function(row, i){
          if(row[iBT] !== tbt) return;
          if(tcp != null && (iCP == null || row[iCP] !== tcp)) return;
          var feature = fillTemplate(t.fmt, row, hi, iQtyName, iQty);
          var qtyVal = (iQty != null ? row[iQty] : "");
          bill.push({
            code: "",
            name: (iCP != null ? (row[iCP]==null ? "" : row[iCP]) : ""),
            feature: feature,
            unit: (iUnit != null ? (row[iUnit]==null ? "" : row[iUnit]) : ""),
            qty: qtyVal,
            refs: [{ sheet: sheet.name, rowIdx: i+2, qCol: iQty }]
          });
        });
      });
    });
    return bill;
  }

  // 按「项目特征」完全相同的行自动汇总：工程量求和，反查收集所有来源行
  function mergeBill(bill){
    var map = {}, order = [];
    bill.forEach(function(b){
      var key = String(b.feature);
      if(Object.prototype.hasOwnProperty.call(map, key)){
        var m = map[key];
        var q1 = parseFloat(m.qty), q2 = parseFloat(b.qty);
        m.qty = (isNaN(q1) ? 0 : q1) + (isNaN(q2) ? 0 : q2);
        if(m.qty != null) m.qty = Math.round(m.qty * 1000) / 1000; // 去浮点误差，保留3位
        b.refs.forEach(function(rf){ m.refs.push(rf); });
      } else {
        var nb = { code:b.code, name:b.name, feature:b.feature, unit:b.unit, qty:b.qty, refs:b.refs.slice() };
        map[key] = nb;
        order.push(key);
      }
    });
    return order.map(function(k){ return map[k]; });
  }

  function renderBill(){
    var q = $("bill-search").value.trim().toLowerCase();
    var html = "<table><thead><tr><th>项目编码</th><th>项目名称</th><th>项目特征</th><th>单位</th><th>工程量</th><th>反查</th></tr></thead><tbody>";
    BILL.forEach(function(b, i){
      var hay = (b.code + " " + b.name + " " + b.feature + " " + b.unit + " " + b.qty).toLowerCase();
      if(q && hay.indexOf(q) < 0) return;
      var cnt = b.refs.length;
      var label = cnt > 1 ? ("反查 (" + cnt + ")") : "反查";
      html += "<tr><td>" + esc(b.code) + "</td><td>" + esc(b.name) + "</td><td class='feature'>" +
        esc(b.feature).replace(/\n/g, "<br>") + "</td><td>" + esc(b.unit) + "</td><td>" + esc(b.qty) +
        "</td><td class='refs'><button class='ref-btn' data-i='" + i + "'>" + label + "</button></td></tr>";
    });
    html += "</tbody></table>";
    $("bill-table").innerHTML = html;
    Array.prototype.forEach.call(document.querySelectorAll(".ref-btn"), function(btn){
      btn.addEventListener("click", function(){
        var i = parseInt(btn.getAttribute("data-i"), 10);
        openRefModal(BILL[i]);
      });
    });
  }

  // 反查：多条来源时弹窗列出全部，可逐条定位
  function openRefModal(b){
    var box = $("ref-modal-list");
    if(!box) return;
    box.innerHTML = "";
    b.refs.forEach(function(rf, j){
      var loc = rf.qCol != null ? (colLetter(rf.qCol+1) + rf.rowIdx) : (rf.rowIdx + "");
      var li = document.createElement("div");
      li.className = "ref-item";
      li.innerHTML = "<span class='ref-idx'>" + (j+1) + "</span>" +
        "<span class='ref-loc'>" + esc(rf.sheet) + "!" + esc(loc) + "</span>";
      li.addEventListener("click", function(){
        closeRefModal();
        gotoRef(rf);
      });
      box.appendChild(li);
    });
    $("ref-modal-title").textContent = "反查来源（共 " + b.refs.length + " 处）";
    $("ref-modal").style.display = "flex";
  }
  function closeRefModal(){
    var m = $("ref-modal");
    if(m) m.style.display = "none";
  }

  function gotoRef(ref){
    showSec("sec-clean");
    var sel = $("clean-sheet");
    var found = false;
    Array.prototype.forEach.call(sel.options, function(o){
      var s = CLEANED[parseInt(o.value, 10)];
      if(s && s.name === ref.sheet){ sel.value = o.value; found = true; }
    });
    if(found) renderClean();
    var tr = document.querySelector("#clean-table tr[data-row='" + ref.rowIdx + "']");
    if(tr){
      tr.classList.add("hl");
      tr.scrollIntoView({block:"center"});
      setTimeout(function(){ tr.classList.remove("hl"); }, 2500);
    }
  }

  function exportClean(){
    if(!CLEANED.length){ alert("暂无清洗数据"); return; }
    var wb = XLSX.utils.book_new();
    CLEANED.forEach(function(sheet){
      var aoa = [sheet.headers].concat(sheet.rows.map(function(r){
        return r.map(function(x){ return x==null ? "" : x; });
      }));
      var ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name.slice(0,31));
    });
    XLSX.writeFile(wb, "清洗数据表.xlsx");
  }

  function exportBill(){
    if(!BILL.length){ alert("暂无清单"); return; }
    var wb = XLSX.utils.book_new();

    // 样式 helper（xlsx-js-style 1.2.0 API：cell.s）
    var THIN = { style:"thin", color:{ rgb:"FF000000" } };
    var BORDER = { top:THIN, left:THIN, bottom:THIN, right:THIN };
    function styleHeader(){
      return { font:{ bold:true, sz:11, name:"微软雅黑" },
               fill:{ fgColor:{ rgb:"FFD9E1F2" } },
               alignment:{ horizontal:"center", vertical:"center", wrapText:true },
               border: BORDER };
    }
    function styleFeature(){
      return { font:{ sz:10, name:"微软雅黑" },
               alignment:{ horizontal:"left", vertical:"top", wrapText:true },
               border: BORDER };
    }
    function styleTextBottom(){
      return { font:{ sz:10, name:"微软雅黑" },
               alignment:{ horizontal:"left", vertical:"bottom", wrapText:true },
               border: BORDER };
    }
    function styleQty(){
      return { font:{ sz:10, name:"微软雅黑" },
               alignment:{ horizontal:"right", vertical:"bottom" },
               border: BORDER };
    }
    function styleCode(){
      return { font:{ sz:10, name:"微软雅黑" },
               alignment:{ horizontal:"center", vertical:"bottom" },
               border: BORDER };
    }
    function styleRef(){
      return { font:{ sz:10, name:"微软雅黑", color:{ rgb:"FF0563C1" }, underline:true },
               alignment:{ horizontal:"left", vertical:"bottom", wrapText:true },
               border: BORDER };
    }

    // 1) 项目特征表 —— 放第一个工作表
    var aoa = [["项目编码","项目名称","项目特征","单位","工程量","反查"]];
    BILL.forEach(function(b){
      var refTexts = b.refs.map(function(rf){
        var loc = rf.qCol != null ? (colLetter(rf.qCol+1) + rf.rowIdx) : (rf.rowIdx + "");
        return rf.sheet + "!" + loc;
      });
      aoa.push([b.code, b.name, b.feature, b.unit, b.qty, refTexts.join("\n")]);
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);

    // 列宽：A编码窄 / B名称中 / C特征宽 / D单位窄 / E工程量中 / F反查中
    ws["!cols"] = [
      { wch: 10 },  // 项目编码
      { wch: 14 },  // 项目名称
      { wch: 55 },  // 项目特征
      { wch: 6  },  // 单位
      { wch: 12 },  // 工程量
      { wch: 26 }   // 反查
    ];

    // 行高：表头 22pt，数据行 50pt（容纳 3 行 14 号字符 + 内边距）
    ws["!rows"] = [{ hpt: 22 }];
    for(var r=1; r<aoa.length; r++) ws["!rows"].push({ hpt: 50 });

    // 表头样式
    for(var c=0; c<6; c++){
      var a = XLSX.utils.encode_cell({ r:0, c:c });
      if(ws[a]) ws[a].s = styleHeader();
    }
    // 数据行样式
    for(var i=0; i<BILL.length; i++){
      var rr = i + 1;
      [0,1,2,3,4,5].forEach(function(c){
        var a = XLSX.utils.encode_cell({ r:rr, c:c });
        if(!ws[a]) return;
        if(c === 0)      ws[a].s = styleCode();
        else if(c === 1) ws[a].s = styleTextBottom();
        else if(c === 2) ws[a].s = styleFeature();
        else if(c === 3) ws[a].s = styleTextBottom();
        else if(c === 4) ws[a].s = styleQty();
        else if(c === 5) ws[a].s = styleRef();
      });
    }

    // 反查超链接（指向第一条来源）
    BILL.forEach(function(b, i){
      var rf = b.refs[0];
      if(!rf || rf.qCol == null) return;
      var cellAddr = colLetter(6) + (i+2);
      var tgt = colLetter(rf.qCol+1) + rf.rowIdx;
      if(ws[cellAddr]){
        ws[cellAddr].l = { Target: "#'" + rf.sheet + "'!" + tgt, Tooltip: "反查清洗数据（共 " + b.refs.length + " 处）" };
      }
    });

    // 冻结首行首两列
    ws["!freeze"] = { xSplit: 2, ySplit: 1 };
    ws["!views"] = [{ state:"frozen", xSplit:2, ySplit:1, topLeftCell:"C2", activePane:"bottomRight" }];

    XLSX.utils.book_append_sheet(wb, ws, "项目特征");

    // 2) 清洗数据表（11 张）放后面，供反查跳转
    CLEANED.forEach(function(sheet){
      var aoa2 = [sheet.headers].concat(sheet.rows.map(function(r){
        return r.map(function(x){ return x==null ? "" : x; });
      }));
      var ws2 = XLSX.utils.aoa_to_sheet(aoa2);
      ws2["!cols"] = sheet.headers.map(function(h){
        var w = String(h||"").length * 1.6 + 4;
        return { wch: Math.max(8, Math.min(w, 32)) };
      });
      var addr = XLSX.utils.encode_cell({ r:0, c:0 });
      if(ws2[addr]) ws2[addr].s = styleHeader();
      sheet.headers.forEach(function(_, c){
        if(c===0) return;
        var a = XLSX.utils.encode_cell({ r:0, c:c });
        if(ws2[a]) ws2[a].s = styleHeader();
      });
      ws2["!rows"] = [{ hpt: 22 }];
      XLSX.utils.book_append_sheet(wb, ws2, sheet.name.slice(0,31));
    });

    XLSX.writeFile(wb, "鹏宇项目项目特征.xlsx");
  }

  function init(){
    initNav();
    $("file-summary").addEventListener("change", function(e){
      if(e.target.files[0]) handleSummaryFile(e.target.files[0]);
    });
    $("file-tpl").addEventListener("change", function(e){
      if(e.target.files[0]) handleTplFile(e.target.files[0]);
    });
    $("btn-clean").addEventListener("click", function(){
      if(!RAW){ alert("请先导入汇总表"); return; }
      CLEANED = cleanWorkbook(RAW);
      var sel = $("clean-sheet"); sel.innerHTML = "";
      CLEANED.forEach(function(s, idx){
        var o = document.createElement("option"); o.value = idx; o.textContent = s.name; sel.appendChild(o);
      });
      renderClean();
      $("clean-status").textContent = "清洗完成：共 " + CLEANED.length + " 张表，合计 " +
        CLEANED.reduce(function(a, s){ return a + s.rows.length; }, 0) + " 行数据。";
      showSec("sec-clean");
    });
    $("clean-sheet").addEventListener("change", renderClean);
    $("clean-search").addEventListener("input", renderClean);
    $("btn-export-clean").addEventListener("click", exportClean);
    $("btn-gen").addEventListener("click", function(){
      if(!CLEANED.length){ alert("请先清洗数据"); return; }
      BILL = mergeBill(genBill());
      renderBill();
      $("bill-status").textContent = "已生成清单 " + BILL.length + " 行。";
      showSec("sec-bill");
    });
    $("bill-search").addEventListener("input", renderBill);
    $("btn-export-bill").addEventListener("click", exportBill);

    // 项目特征模板模块
    $("btn-tpl-add").addEventListener("click", function(){
      TEMPLATES.push({cat:"", fmt:"", qtyNameMap:"", qtyMap:""}); renderTpl(); onTemplatesChanged();
    });
    $("btn-export-tpl").addEventListener("click", exportTpl);
    $("btn-tpl-reset").addEventListener("click", function(){
      TEMPLATES = DEFAULT_TPL.map(function(x){ return {cat:x.cat, fmt:x.fmt, qtyNameMap:"", qtyMap:""}; });
      renderTpl(); onTemplatesChanged(); $("tpl-status").textContent = "已恢复内置默认模板（电气管线：电缆 / 电线 / 配管 / 桥架）。";
    });
    var tplBox = $("tpl-table");
    tplBox.addEventListener("input", function(e){
      var el = e.target;
      var i = parseInt(el.getAttribute("data-i"), 10);
      if(isNaN(i) || !TEMPLATES[i]) return;
      if(el.classList.contains("tpl-cat")) TEMPLATES[i].cat = el.value;
      else if(el.classList.contains("tpl-fmt")) TEMPLATES[i].fmt = el.value;
      else if(el.classList.contains("tpl-qtyname")) TEMPLATES[i].qtyNameMap = el.value;
      else if(el.classList.contains("tpl-qty")) TEMPLATES[i].qtyMap = el.value;
      onTemplatesChanged();
    });
    tplBox.addEventListener("click", function(e){
      if(e.target.classList.contains("tpl-del")){
        var i = parseInt(e.target.getAttribute("data-i"), 10);
        if(!isNaN(i)){ TEMPLATES.splice(i, 1); renderTpl(); onTemplatesChanged(); }
      }
    });
    renderTpl();

    // 反查弹窗关闭
    $("btn-ref-modal-close").addEventListener("click", closeRefModal);
    var rm = $("ref-modal");
    if(rm) rm.addEventListener("click", function(e){ if(e.target === rm) closeRefModal(); });
    document.addEventListener("keydown", function(e){ if(e.key === "Escape") closeRefModal(); });

    showSec("sec-import");
  }

  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded", init);
  } else { init(); }
})();
