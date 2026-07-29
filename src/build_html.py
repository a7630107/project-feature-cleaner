import re

with open(r"F:\项目特征(AI)\_sheetjs.js", "r", encoding="utf-8") as f:
    sheetjs = f.read()
# 防止内联脚本里出现 </script> 截断
sheetjs = re.sub(r"(?i)</script", "<\\/script", sheetjs)

with open(r"F:\项目特征(AI)\app.js", "r", encoding="utf-8") as f:
    appjs = f.read()
appjs = re.sub(r"(?i)</script", "<\\/script", appjs)

HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>项目特征清洗与清单生成工具</title>
<style>
* { box-sizing: border-box; }
body { margin:0; font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif; color:#222; background:#f5f6f8; }
#sidebar { position:fixed; top:0; left:0; bottom:0; width:210px; background:#1f2a44; color:#fff; padding-top:10px; transition:width .2s; overflow:hidden; z-index:10; }
#sidebar.collapsed { width:60px; }
#sidebar-toggle { background:none; border:none; color:#fff; font-size:20px; cursor:pointer; padding:8px 16px; }
.nav { margin-top:10px; }
.nav-item { display:flex; align-items:center; padding:12px 16px; cursor:pointer; white-space:nowrap; }
.nav-item:hover { background:#2c3a5e; }
.nav-item.active { background:#3b82f6; }
.nav-item .ico { font-size:18px; width:28px; text-align:center; }
.nav-item .label { margin-left:10px; }
#sidebar.collapsed .label { display:none; }
#main { margin-left:210px; padding:20px 28px; transition:margin-left .2s; }
#sidebar.collapsed ~ #main { margin-left:60px; }
header h1 { font-size:20px; margin:0 0 16px; }
h2 { font-size:17px; border-left:4px solid #3b82f6; padding-left:10px; margin-top:0; }
h3 { font-size:15px; margin:18px 0 6px; }
p { line-height:1.7; color:#444; }
.row { margin:10px 0; }
.btn-file { display:inline-block; background:#3b82f6; color:#fff; padding:8px 14px; border-radius:6px; cursor:pointer; }
.btn-file input { display:none; }
button { font-size:14px; padding:8px 14px; border:none; border-radius:6px; background:#e2e8f0; cursor:pointer; }
button.primary { background:#3b82f6; color:#fff; }
button:disabled { opacity:.5; cursor:not-allowed; }
.status { background:#eef2ff; border:1px solid #c7d2fe; padding:10px; border-radius:6px; margin:10px 0; font-size:13px; line-height:1.6; }
.toolbar { display:flex; gap:10px; align-items:center; margin:12px 0; flex-wrap:wrap; }
.toolbar input[type=text] { padding:7px 10px; border:1px solid #cbd5e1; border-radius:6px; width:220px; }
.toolbar select { padding:7px 10px; border:1px solid #cbd5e1; border-radius:6px; }
.table-wrap { max-height:70vh; overflow:auto; border:1px solid #d8dee9; background:#fff; border-radius:6px; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th,td { border:1px solid #e2e8f0; padding:6px 8px; text-align:left; vertical-align:top; white-space:nowrap; }
th { background:#f1f5f9; position:sticky; top:0; z-index:1; }
td.feature { white-space:pre-wrap; min-width:280px; }
tr.hl { background:#ffe58f !important; }
.ref-btn { background:#10b981; color:#fff; padding:4px 8px; font-size:12px; margin:2px 3px 2px 0; }
td.refs { white-space:normal; }
.merge-note { display:inline-block; background:#fde68a; color:#92400e; font-size:11px; padding:1px 6px; border-radius:4px; margin:0 2px 3px 0; }
.ref-btn { background:#10b981; color:#fff; padding:4px 10px; font-size:12px; }
/* 反查弹窗 */
#ref-modal { position:fixed; inset:0; background:rgba(15,23,42,.45); display:none; align-items:center; justify-content:center; z-index:100; }
#ref-modal .modal-box { background:#fff; width:420px; max-width:90vw; border-radius:10px; box-shadow:0 10px 40px rgba(0,0,0,.25); overflow:hidden; }
#ref-modal .modal-head { display:flex; justify-content:space-between; align-items:center; padding:12px 16px; background:#1f2a44; color:#fff; }
#ref-modal .modal-head h3 { margin:0; font-size:15px; border:none; padding:0; color:#fff; }
#ref-modal .modal-close { background:none; border:none; color:#fff; font-size:20px; cursor:pointer; line-height:1; }
#ref-modal .modal-body { padding:8px 0; max-height:60vh; overflow:auto; }
#ref-modal .ref-item { display:flex; align-items:center; gap:10px; padding:10px 16px; cursor:pointer; border-bottom:1px solid #f0f0f0; font-size:13px; }
#ref-modal .ref-item:hover { background:#eef2ff; }
#ref-modal .ref-idx { background:#3b82f6; color:#fff; border-radius:50%; width:20px; height:20px; line-height:20px; text-align:center; font-size:11px; flex:0 0 auto; }
#ref-modal .ref-loc { font-family:"Courier New",monospace; }
#tpl-table textarea { width:100%; min-height:64px; font-family:inherit; font-size:12px; border:1px solid #cbd5e1; border-radius:4px; padding:4px 6px; resize:vertical; line-height:1.5; }
#tpl-table input.tpl-cat { width:200px; padding:5px 6px; border:1px solid #cbd5e1; border-radius:4px; font-size:13px; }
    #tpl-table input.tpl-qty { width:140px; padding:5px 6px; border:1px solid #cbd5e1; border-radius:4px; font-size:13px; }
    #tpl-table input.tpl-qtyname { width:140px; padding:5px 6px; border:1px solid #cbd5e1; border-radius:4px; font-size:13px; }
    #tpl-table .tpl-hint { display:block; font-size:11px; color:#64748b; font-weight:normal; line-height:1.4; margin-top:2px; }
#tpl-table .tpl-del { background:#ef4444; color:#fff; padding:4px 10px; font-size:12px; }
footer { margin-top:24px; color:#888; font-size:12px; }
</style>
</head>
<body>
  <div id="sidebar">
    <button id="sidebar-toggle" title="折叠/展开">☰</button>
    <div class="nav">
      <div class="nav-item active" data-sec="sec-import"><span class="ico">📥</span><span class="label">导入汇总表</span></div>
      <div class="nav-item" data-sec="sec-clean"><span class="ico">🧹</span><span class="label">清洗数据表</span></div>
      <div class="nav-item" data-sec="sec-bill"><span class="ico">📋</span><span class="label">清单表</span></div>
      <div class="nav-item" data-sec="sec-tpl"><span class="ico">📐</span><span class="label">项目特征模板</span></div>
    </div>
  </div>
  <div id="main">
    <header><h1>项目特征清洗与清单生成工具</h1></header>

    <section id="sec-import">
      <h2>① 导入汇总表</h2>
      <p>选择本地的汇总表（.xlsx / .xls / .csv），系统先解析并预览，再点「数据清洗」生成清洗后的数据。</p>
      <div class="row">
        <a class="btn-file" href="GQI2021安装报表设置.BSZ" download="GQI2021安装报表设置.BSZ" title="下载后在广联达GQI2021软件中导入此报表设置，再导出报表即可被本工具识别">📥 下载广联达GQI2021报表设置</a>
      </div>
      <p style="font-size:13px;color:#555;margin:-4px 0 4px">如使用广联达GQI2021导出报表，请先下载报表设置模板并在GQI中导入；导出后再用下方「选择汇总表」导入。</p>
      <div class="row">
        <label class="btn-file">选择汇总表<input type="file" id="file-summary" accept=".xlsx,.xls,.csv"></label>
      </div>
      <div id="import-status" class="status">尚未导入文件。</div>
      <button id="btn-clean" class="primary" disabled>数据清洗</button>
    </section>

    <section id="sec-clean" style="display:none">
      <h2>② 清洗数据表</h2>
      <div class="toolbar">
        <label>工作表：<select id="clean-sheet"></select></label>
        <input type="text" id="clean-search" placeholder="搜索 / 筛选…">
        <button id="btn-export-clean">导出 Excel</button>
        <button id="btn-gen" class="primary">一键生成清单</button>
      </div>
      <div id="clean-status" class="status"></div>
      <div id="clean-table" class="table-wrap"></div>
    </section>

    <section id="sec-bill" style="display:none">
      <h2>③ 清单表</h2>
      <div class="toolbar">
        <input type="text" id="bill-search" placeholder="搜索 / 筛选…">
        <button id="btn-export-bill">导出 Excel</button>
      </div>
      <div id="bill-status" class="status"></div>
      <div id="bill-table" class="table-wrap"></div>
    </section>

    <section id="sec-tpl" style="display:none">
      <h2>④ 项目特征模板</h2>
      <p>模板库用于把清洗数据自动套成清单的「项目特征」。每行一条，含 <b>分类</b>、<b>项目特征格式</b>、<b>工程量名称映射</b>、<b>工程量映射</b> 四列。
         「分类」形如 <code>构建类型~计算项目</code>（~ 分隔，只写构建类型则匹配该表全部）；「项目特征格式」中用 <code>{列名}</code> 引用清洗表的表头（如 <code>{材质}</code>、<code>{规格型号}</code>）。
         工程量名称映射默认对应 <code>工程量名称</code> 列，工程量映射默认对应 <code>工程量</code> 列；在格式中分别用占位符 <code>{对应工程量名称表头名}</code> 和 <code>{对应工程量表头名}</code> 引用（映射列填 <code>{表头名}</code> 可改用其他列）。可直接在表格里修改，也可导入 / 导出 Excel 模板库。</p>
      <div class="toolbar">
        <button id="btn-tpl-add">新增模板</button>
        <label class="btn-file">导入模板库<input type="file" id="file-tpl" accept=".xlsx,.xls,.csv"></label>
        <button id="btn-export-tpl">导出模板库</button>
        <button id="btn-tpl-reset">恢复默认</button>
      </div>
      <div id="tpl-status" class="status">内置默认模板（电气管线：电缆 / 电线 / 配管 / 桥架）。</div>
      <div id="tpl-table" class="table-wrap"></div>
    </section>

    <footer>离线单文件工具 · 数据仅在本地浏览器处理，不会上传。</footer>
  </div>

  <!-- 反查来源弹窗 -->
  <div id="ref-modal">
    <div class="modal-box">
      <div class="modal-head">
        <h3 id="ref-modal-title">反查来源</h3>
        <button class="modal-close" id="btn-ref-modal-close" title="关闭">×</button>
      </div>
      <div class="modal-body" id="ref-modal-list"></div>
    </div>
  </div>

<script>
__SHEETJS__
</script>
<script>
__APPJS__
</script>
</body>
</html>
"""

out = HTML.replace("__SHEETJS__", sheetjs).replace("__APPJS__", appjs)
dst = r"F:\项目特征(AI)\项目特征清洗与清单生成工具.html"
with open(dst, "w", encoding="utf-8") as f:
    f.write(out)
print("written:", dst, "size:", len(out))
