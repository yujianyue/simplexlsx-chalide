<?php
/**
 * SimpleXLSXPro 在线转换演示台
 * ------------------------------------------------------------------
 * 单文件应用：既是页面，也是 AJAX 接口，无需框架、无需数据库。
 *
 * 功能：
 *   1. 自动扫描 demo/ 目录下所有 .xlsx，按修改时间降序作为下拉选项
 *   2. 支持输出 HTML（保留合并单元格 + 图片）/ TSV / CSV / JSON
 *   3. AJAX 提交、AJAX 返回结果（含耗时、行列数、图片数等元信息）
 *   4. 前端提供「复制」与「下载」两个动作
 *
 * 接口（ajax=1 时返回 JSON，dl=1 时直接下载附件）：
 *   POST demo.php   ajax=1&action=convert&file=…&format=…
 *   POST demo.php   ajax=1&action=info&file=…
 *   GET  demo.php   dl=1&file=…&format=…           直接下载
 *
 * 安全：文件名只取 basename，并用 realpath 校验必须落在 demo/ 目录内，
 *       杜绝 ../../ 路径穿越。
 */

declare(strict_types=1);

// AJAX 接口必须返回干净的 JSON：关掉错误直显，改写入 error_log
ini_set('display_errors', '0');
ini_set('log_errors', '1');

chdir(__DIR__);                       // 让 Img 等相对路径以本文件所在目录为基准
require_once __DIR__ . '/src/SimpleXLSXPro.php';

use Shuchkin\SimpleXLSXPro;

const DEMO_DIR = 'demo';              // 扫描目录（相对本文件）
const IMG_DIR  = 'demo/Img';          // 图片输出目录，同时是可直接访问的 URL 路径
const MAX_B64  = 5242880;             // 超过 5MB 不再随 JSON 回传（改用 dl=1 下载）

// ------------------------------------------------------- 依赖自检（版本兼容）
// 演示台依赖 src/SimpleXLSXPro.php 中的增强方法。若只更新了 demo.php 而
// SimpleXLSXPro.php 仍是旧版，过去会直接 Fatal：Call to undefined method …
// 这里改成启动时自检 + 运行时降级，缺方法只提示，不再白屏。

/** 必须的增强方法（缺了演示台无法工作） */
const REQUIRED_METHODS = [
    'rowsPro', 'mergeCells', 'calculatedCells',
    'extractImages', 'imageList',
    'toCSV', 'toTSV', 'toJSON', 'toHTMLMerged',
];

/** 可选增强方法（缺了自动降级，不影响主流程） */
const OPTIONAL_METHODS = ['setImagePrefix', 'mergeAt', 'imageMap', 'saveJSON'];

/**
 * 返回缺失的方法清单：['required'=>[…], 'optional'=>[…]]
 */
function selfCheck(): array
{
    $miss = ['required' => [], 'optional' => []];
    foreach (REQUIRED_METHODS as $m) {
        if (!method_exists(SimpleXLSXPro::class, $m)) {
            $miss['required'][] = $m;
        }
    }
    foreach (OPTIONAL_METHODS as $m) {
        if (!method_exists(SimpleXLSXPro::class, $m)) {
            $miss['optional'][] = $m;
        }
    }

    return $miss;
}

/**
 * 设置图片文件名前缀；旧版类没有 setImagePrefix() 时静默跳过，
 * 调用方会改用「每份 xlsx 一个子目录」的方式隔离图片。
 */
function trySetImagePrefix(SimpleXLSXPro $x, string $prefix): bool
{
    if (method_exists($x, 'setImagePrefix')) {
        $x->setImagePrefix($prefix);

        return true;
    }

    return false;
}

/**
 * 当前加载的 SimpleXLSXPro 版本标识（便于排查「类文件没替换」的问题）
 */
function libVersion(): string
{
    $v = defined(Shuchkin\SimpleXLSXPro::class . '::PRO_VERSION')
        ? (string)constant(Shuchkin\SimpleXLSXPro::class . '::PRO_VERSION')
        : '未知（1.3.0 以前）';
    $f = __DIR__ . '/src/SimpleXLSXPro.php';

    return $v . ' · ' . date('Y-m-d H:i', (int)@filemtime($f));
}

// ---------------------------------------------------------------- 工具函数

function h($s): string
{
    return htmlspecialchars((string)$s, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

/**
 * 解析出安全的 xlsx 绝对路径，非法则返回 ''
 */
function resolveFile($name): string
{
    $name = basename(trim((string)$name));
    if ($name === '' || !preg_match('/\.xlsx$/i', $name)) {
        return '';
    }
    $base = realpath(__DIR__ . '/' . DEMO_DIR);
    $path = realpath($base . '/' . $name);
    if ($base === false || $path === false || !is_file($path) || dirname($path) !== $base) {
        return '';
    }

    return $path;
}

/**
 * demo 目录下的 xlsx 列表，按修改时间降序
 */
function listFiles(): array
{
    $base = __DIR__ . '/' . DEMO_DIR;
    if (!is_dir($base)) {
        return [];
    }
    $out = [];
    foreach (glob($base . '/*.xlsx') ?: [] as $p) {
        if (!is_file($p)) {
            continue;
        }
        $out[] = [
            'name' => basename($p),
            'mtime' => filemtime($p),
            'time' => date('Y-m-d H:i:s', (int)filemtime($p)),
            'size' => filesize($p),
        ];
    }
    usort($out, static function ($a, $b) {
        return $b['mtime'] <=> $a['mtime'];   // 时间降序：最近的排最前
    });

    return $out;
}

function humanSize(int $b): string
{
    foreach (['B', 'KB', 'MB', 'GB'] as $i => $u) {
        if ($b < 1024 || $i === 3) {
            return ($i === 0 ? $b : round($b / 1024 ** $i, 1)) . ' ' . $u;
        }
    }

    return $b . ' B';
}

function jsonOut(array $data, int $code = 200): void
{
    http_response_code($code);
    header('Content-Type: application/json; charset=UTF-8');
    header('Cache-Control: no-store, no-cache, must-revalidate');
    echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit;
}

/**
 * 统一的输入读取（FormData / 普通表单都走 $_POST）
 */
function param(string $k, $default = '')
{
    return isset($_POST[$k]) ? $_POST[$k] : (isset($_GET[$k]) ? $_GET[$k] : $default);
}

const FORMATS = [
    'html' => ['ext' => 'html', 'mime' => 'text/html; charset=UTF-8',      'label' => 'HTML（保留合并 + 图片）'],
    'csv'  => ['ext' => 'csv',  'mime' => 'text/csv; charset=UTF-8',       'label' => 'CSV（逗号分隔）'],
    'tsv'  => ['ext' => 'tsv',  'mime' => 'text/tab-separated-values; charset=UTF-8', 'label' => 'TSV（制表符分隔）'],
    'json' => ['ext' => 'json', 'mime' => 'application/json; charset=UTF-8', 'label' => 'JSON（结构化）'],
];

// ---------------------------------------------------------------- 执行转换

/**
 * 把 <table> 片段包成可直接打开 / 预览的完整 HTML 文档
 */
function buildHtmlDoc(string $table, string $title, string $charset = 'UTF-8'): string
{
    return '<!DOCTYPE html>' . "\n"
        . '<html lang="zh-CN">' . "\n<head>\n"
        . '<meta charset="' . htmlspecialchars($charset, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '">' . "\n"
        . '<meta name="viewport" content="width=device-width, initial-scale=1">' . "\n"
        . '<title>' . htmlspecialchars($title, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '</title>' . "\n"
        . "<style>\n"
        . "body{font:14px/1.6 -apple-system,'Microsoft YaHei',sans-serif;margin:20px;color:#1f2937}\n"
        . "table.excel{border-collapse:collapse}\n"
        . "table.excel td,table.excel th{border:1px solid #cbd5e1;padding:4px 10px;vertical-align:middle}\n"
        . "table.excel th{background:#eef2ff}\n"
        . "table.excel img{vertical-align:middle}\n"
        . "</style>\n</head>\n<body>\n" . $table . "\n</body>\n</html>";
}

/**
 * @return array  ['bytes'=>string, 'meta'=>array]
 */
function doConvert(string $file, string $format, array $o): array
{
    $xlsx = SimpleXLSXPro::parse($file);
    if (!$xlsx) {
        throw new RuntimeException('解析失败：' . SimpleXLSXPro::parseError());
    }

    $sheets = $xlsx->sheetNames() ?: [];
    $sheet  = (int)$o['sheet'];
    if ($sheet < 0 || $sheet >= count($sheets)) {
        $sheet = 0;
    }

    // 图片：提取到 demo/Img，并记录相对路径（可直接当 URL 用）
    $images = [];
    if (!empty($o['extract_img'])) {
        // 多份 xlsx 共用一个 Img 目录时，用文件名做前缀避免同名图片互相覆盖
        $stem   = pathinfo($file, PATHINFO_FILENAME);
        $hasSet = trySetImagePrefix($xlsx, $stem . '_');
        // 旧版类没有 setImagePrefix()：退化为「每份 xlsx 一个子目录」，同样隔离
        $imgDir = $hasSet ? IMG_DIR : IMG_DIR . '/' . $stem;
        if (!is_dir($imgDir) && !@mkdir($imgDir, 0777, true) && !is_dir($imgDir)) {
            $imgDir = IMG_DIR;
        }
        $xlsx->extractImages($imgDir, true);
        foreach ($xlsx->imageList() as $im) {
            $images[] = [
                'cell' => $im['cell'] ?? '',
                'path' => $im['path'] ?? '',
                'name' => $im['name'] ?? basename((string)($im['path'] ?? '')),
            ];
        }
    }

    switch ($format) {
        case 'html':
            // 库里 toHTMLMerged 只输出 <table>（full/title/charset 选项会被忽略），
            // 这里统一先生成 UTF-8 片段，再由下方 wrapper 组装完整文档；
            // 若用户选择非 UTF-8 编码，最后做一次字节级转码并同步 <meta charset>，
            // 保证下载的 HTML 不出现乱码 / 半字符损坏。
            $table = $xlsx->toHTMLMerged($sheet, [
                'images'     => !empty($o['extract_img']),
                'head'       => true,
                'class'      => 'excel',
                'img_height' => 60,
            ]);
            $bytes = buildHtmlDoc($table, basename($file) . ' — ' . ($sheets[$sheet] ?? 'Sheet1'), 'UTF-8');
            if ($o['encoding'] !== 'UTF-8') {
                $conv = @mb_convert_encoding($bytes, $o['encoding'], 'UTF-8');
                if (is_string($conv) && $conv !== '') {
                    $bytes = str_replace('<meta charset="UTF-8">',
                        '<meta charset="' . htmlspecialchars($o['encoding'], ENT_QUOTES, 'UTF-8') . '">', $conv);
                }
            }
            break;

        case 'csv':
            $bytes = $xlsx->toCSV($sheet, !empty($o['fill_merged']), !empty($o['bom']), $o['encoding']);
            break;

        case 'tsv':
            $bytes = $xlsx->toTSV($sheet, !empty($o['fill_merged']), !empty($o['bom']), $o['encoding']);
            break;

        case 'json':
        default:
            $format = 'json';
            $bytes = $xlsx->toJSON($sheet, [
                'fill_merged' => !empty($o['fill_merged']),
                'header'      => (int)$o['header'],   // 0=不用表头；1=第一行；>1=指定行号
                'cells'       => !empty($o['cells']),
                'meta'        => false,
                'pretty'      => true,
                'encoding'    => $o['encoding'],
            ]);
            break;
    }

    // 统计信息
    $dim = $xlsx->dimension($sheet) ?: [0, 0];

    return [
        'bytes' => (string)$bytes,
        'meta'  => [
            'sheet'        => $sheet,
            'sheet_name'   => $sheets[$sheet] ?? 'Sheet1',
            'sheets'       => $sheets,
            'rows'         => (int)($dim[1] ?? 0),
            'cols'         => (int)($dim[0] ?? 0),
            'merges'       => count($xlsx->mergeCells($sheet) ?: []),
            'formulas'     => count($xlsx->calculatedCells($sheet) ?: []),
            'images'       => count($images),
            'image_list'   => $images,
            'fill_merged'  => !empty($o['fill_merged']),
            'encoding'     => $o['encoding'],
        ],
    ];
}

function optionsFromRequest(): array
{
    $enc = strtoupper(trim((string)param('encoding', 'UTF-8')));
    if (!in_array($enc, ['UTF-8', 'GBK', 'GB18030', 'BIG5'], true)) {
        $enc = 'UTF-8';
    }

    return [
        'sheet'       => (int)param('sheet', 0),
        'fill_merged' => param('fill_merged', '1'),
        'extract_img' => param('extract_img', '1'),
        'bom'         => param('bom', '1'),
        'header'      => (int)param('header', 0),
        'cells'       => param('cells', ''),
        'encoding'    => $enc,
    ];
}

// ---------------------------------------------------------------- AJAX 接口

$isAjax = (($_SERVER['HTTP_X_REQUESTED_WITH'] ?? '') === 'XMLHttpRequest')
    || param('ajax', '') !== '';
$action = (string)param('action', 'convert');
$dl     = param('dl', '') !== '';

if ($isAjax || $dl) {
    // 依赖自检：src/SimpleXLSXPro.php 版本过旧时给出明确提示，而不是 Fatal
    $miss = selfCheck();
    if ($miss['required']) {
        $msg = 'src/SimpleXLSXPro.php 版本过旧，缺少方法：' . implode('、', $miss['required'])
            . '。请用本套件 src/ 目录下最新的 SimpleXLSXPro.php（含 XLSXFormula.php）覆盖旧文件。';
        if ($dl) {
            http_response_code(500);
            echo $msg;
            exit;
        }
        jsonOut(['ok' => false, 'error' => $msg, 'missing' => $miss], 500);
    }

    $file   = resolveFile((string)param('file', ''));
    $format = strtolower((string)param('format', 'json'));

    if ($file === '') {
        if ($dl) {
            http_response_code(400);
            echo 'Bad file';
            exit;
        }
        jsonOut(['ok' => false, 'error' => '文件不存在或不在 demo/ 目录下'], 400);
    }
    if (!isset(FORMATS[$format])) {
        $format = 'json';
    }

    // ---- action=info：只返回工作表信息（切换文件时用）
    if ($action === 'info' && !$dl) {
        $x = SimpleXLSXPro::parse($file);
        if (!$x) {
            jsonOut(['ok' => false, 'error' => SimpleXLSXPro::parseError()], 400);
        }
        $sheets = $x->sheetNames() ?: [];
        $dims = [];
        foreach ($sheets as $i => $nm) {
            $d = $x->dimension($i) ?: [0, 0];
            $dims[] = ['index' => $i, 'name' => $nm, 'cols' => (int)$d[0], 'rows' => (int)$d[1]];
        }
        jsonOut(['ok' => true, 'file' => basename($file), 'sheets' => $dims]);
    }

    // ---- 转换
    $t0 = microtime(true);
    try {
        $r = doConvert($file, $format, optionsFromRequest());
    } catch (Throwable $e) {
        if ($dl) {
            http_response_code(500);
            echo $e->getMessage();
            exit;
        }
        jsonOut(['ok' => false, 'error' => $e->getMessage()], 500);
    }
    $ms    = round((microtime(true) - $t0) * 1000, 1);
    $bytes = $r['bytes'];
    $meta  = $r['meta'];
    $meta['ms']    = $ms;
    $meta['bytes'] = strlen($bytes);

    // ---- 直接下载（无 JS 也能用 / 大文件走这条路）
    if ($dl) {
        $name = pathinfo($file, PATHINFO_FILENAME) . '.' . FORMATS[$format]['ext'];
        header('Content-Type: ' . FORMATS[$format]['mime']);
        header('Content-Disposition: attachment; filename="' . rawurlencode($name) . '"; filename*=UTF-8\'\'' . rawurlencode($name));
        header('Content-Length: ' . strlen($bytes));
        echo $bytes;
        exit;
    }

    // ---- JSON 回传：raw 字节用 base64 保证 GBK 等编码无损
    // 预览 / 复制统一用 UTF-8 文本；原始字节走 b64，保证 GBK 下载不失真
    $isUtf8 = $meta['encoding'] === 'UTF-8';
    $text   = $bytes;
    if (!$isUtf8) {
        $conv = @mb_convert_encoding($bytes, 'UTF-8', $meta['encoding']);
        $text = is_string($conv) && $conv !== '' ? $conv : $bytes;
    }

    jsonOut([
        'ok'      => true,
        'file'    => basename($file),
        'format'  => $format,
        'ext'     => FORMATS[$format]['ext'],
        'mime'    => FORMATS[$format]['mime'],
        'text'    => $text,                                    // 预览 / 复制用（UTF-8）
        'textCut' => mb_strlen($text) > 20000 ? mb_substr($text, 0, 20000) : null,
        'b64'     => strlen($bytes) <= MAX_B64 ? base64_encode($bytes) : null,  // 下载用（原始字节）
        'tooBig'  => strlen($bytes) > MAX_B64,
        'isUtf8'  => $isUtf8,
        'meta'    => $meta,
    ]);
}

// ---------------------------------------------------------------- 页面渲染

$files = listFiles();
$cur   = $files[0]['name'] ?? '';
?>
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SimpleXLSXPro 在线转换演示台</title>
<style>
    :root {
        --bg: #f4f6fb; --card: #fff; --line: #e3e8f0; --ink: #1f2937; --muted: #6b7280;
        --brand: #2563eb; --brand-d: #1d4ed8; --ok: #059669; --err: #dc2626; --chip: #eef2ff;
    }
    * { box-sizing: border-box; }
    body {
        margin: 0; padding: 24px; background: var(--bg); color: var(--ink);
        font: 14px/1.6 -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
    }
    .wrap { max-width: 1180px; margin: 0 auto; }
    h1 { font-size: 20px; margin: 0 0 4px; }
    .sub { color: var(--muted); font-size: 13px; margin-bottom: 18px; }
    .card {
        background: var(--card); border: 1px solid var(--line); border-radius: 12px;
        padding: 18px 20px; margin-bottom: 16px;
    }
    .card > h2 {
        font-size: 14px; margin: 0 0 14px; padding-bottom: 10px;
        border-bottom: 1px solid var(--line); letter-spacing: .3px;
    }
    label { font-size: 13px; color: #374151; }
    select, input[type=number] {
        font: inherit; padding: 7px 10px; border: 1px solid #cbd5e1; border-radius: 8px;
        background: #fff; color: var(--ink); max-width: 100%;
    }
    select:focus, input:focus { outline: 2px solid #bfdbfe; border-color: var(--brand); }
    .row { display: flex; flex-wrap: wrap; gap: 18px 26px; align-items: flex-end; }
    .fld { display: flex; flex-direction: column; gap: 6px; min-width: 150px; }
    .fld.grow { flex: 1; min-width: 260px; }
    select#file { min-width: 340px; font-variant-numeric: tabular-nums; }

    .fmts { display: flex; gap: 10px; flex-wrap: wrap; }
    .fmt {
        position: relative; border: 1px solid #cbd5e1; border-radius: 10px;
        padding: 8px 16px; cursor: pointer; background: #fff; transition: .15s; user-select: none;
    }
    .fmt input { position: absolute; opacity: 0; pointer-events: none; }
    .fmt:hover { border-color: var(--brand); }
    .fmt.on { background: var(--chip); border-color: var(--brand); color: var(--brand-d); font-weight: 600; }

    .opts { display: flex; flex-wrap: wrap; gap: 10px 22px; }
    .chk { display: inline-flex; align-items: center; gap: 7px; cursor: pointer; font-size: 13px; }
    .chk input { width: 15px; height: 15px; accent-color: var(--brand); cursor: pointer; }

    button {
        font: inherit; cursor: pointer; border-radius: 9px; padding: 9px 20px;
        border: 1px solid var(--brand); background: var(--brand); color: #fff; transition: .15s;
    }
    button:hover { background: var(--brand-d); }
    button.ghost { background: #fff; color: var(--brand); }
    button.ghost:hover { background: var(--chip); }
    button[disabled] { opacity: .55; cursor: not-allowed; }
    .actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }

    .stats { display: flex; flex-wrap: wrap; gap: 8px; }
    .stat {
        background: var(--chip); color: #3730a3; border-radius: 999px;
        padding: 4px 12px; font-size: 12px; font-variant-numeric: tabular-nums;
    }
    .stat b { font-weight: 700; }

    .bar {
        display: flex; justify-content: space-between; align-items: center;
        gap: 12px; flex-wrap: wrap; margin-bottom: 12px;
    }
    .tabs { display: inline-flex; border: 1px solid #cbd5e1; border-radius: 9px; overflow: hidden; }
    .tabs button {
        border: 0; border-radius: 0; background: #fff; color: var(--muted); padding: 7px 16px;
    }
    .tabs button.on { background: var(--brand); color: #fff; }

    pre {
        margin: 0; background: #0f172a; color: #e2e8f0; border-radius: 10px;
        padding: 16px; overflow: auto; max-height: 460px; font: 12.5px/1.65 ui-monospace,
        SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace; white-space: pre;
        tab-size: 4;
    }
    iframe {
        width: 100%; height: 460px; border: 1px solid var(--line);
        border-radius: 10px; background: #fff;
    }
    .hide { display: none !important; }

    .empty { color: var(--muted); text-align: center; padding: 40px 0; }
    .err {
        background: #fef2f2; border: 1px solid #fecaca; color: #991b1b;
        padding: 10px 14px; border-radius: 9px; font-size: 13px;
    }
    .note { color: var(--muted); font-size: 12px; margin-top: 10px; }
    .imgs { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 12px; }
    .imgs figure { margin: 0; border: 1px solid var(--line); border-radius: 8px; padding: 6px; background: #fff; }
    .imgs img { display: block; height: 56px; }
    .imgs figcaption { font-size: 11px; color: var(--muted); margin-top: 4px; text-align: center; }

    #toast {
        position: fixed; left: 50%; bottom: 34px; transform: translateX(-50%) translateY(12px);
        background: #111827; color: #fff; padding: 10px 22px; border-radius: 999px;
        font-size: 13px; opacity: 0; pointer-events: none; transition: .22s; z-index: 99;
    }
    #toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
    #toast.bad { background: var(--err); }
</style>
</head>
<body>
<div class="wrap">
    <h1>SimpleXLSXPro 在线转换演示台</h1>
    <div class="sub">
        读取 <code>demo/</code> 下的 xlsx（按修改时间降序）→ 选择输出格式 → AJAX 转换 → 复制 / 下载
        &nbsp;·&nbsp; SimpleXLSXPro <?= h(libVersion()) ?>
    </div>

    <?php
    $chk = selfCheck();
    if ($chk['required'] || $chk['optional']): ?>
        <div class="card" style="border-color:#f59e0b;background:#fffbeb">
            <h2 style="border-color:#f59e0b;color:#92400e">⚠ 依赖自检未通过</h2>
            <?php if ($chk['required']): ?>
                <p style="margin:0 0 8px">
                    <b>缺少必需方法：</b><code><?= h(implode('、', $chk['required'])) ?></code><br>
                    说明当前 <code>src/SimpleXLSXPro.php</code> 是旧版。请用本套件
                    <code>src/</code> 下的 <code>SimpleXLSXPro.php</code> + <code>XLSXFormula.php</code>
                    覆盖同名文件（不要只替换 demo.php）。
                </p>
            <?php endif; ?>
            <?php if ($chk['optional']): ?>
                <p style="margin:0;color:#92400e">
                    缺少可选方法：<code><?= h(implode('、', $chk['optional'])) ?></code>
                    —— 已自动降级（图片改为「每份 xlsx 一个子目录」隔离），功能不受影响。
                </p>
            <?php endif; ?>
        </div>
    <?php endif; ?>

    <?php if (!$files): ?>
        <div class="card"><div class="empty">
            demo/ 目录下还没有 .xlsx 文件。可先运行：<br>
            <code>php examples/00-smoke_test.php</code> 或把任意 xlsx 放进 demo/ 目录。
        </div></div>
    <?php else: ?>

    <form class="card" id="form">
        <h2>① 选择源文件</h2>
        <div class="row">
            <div class="fld grow">
                <label for="file">xlsx 文件（最近修改的排在最前）</label>
                <select id="file" name="file">
                    <?php foreach ($files as $f): ?>
                        <option value="<?= h($f['name']) ?>">
                            <?= h($f['name']) ?> — <?= h(humanSize((int)$f['size'])) ?> · <?= h($f['time']) ?>
                        </option>
                    <?php endforeach; ?>
                </select>
            </div>
            <div class="fld">
                <label for="sheet">工作表</label>
                <select id="sheet" name="sheet"><option value="0">Sheet1</option></select>
            </div>
        </div>
    </form>

    <form class="card">
        <h2>② 输出格式</h2>
        <div class="fmts" id="fmts">
            <?php foreach (FORMATS as $k => $v): ?>
                <label class="fmt <?= $k === 'json' ? 'on' : '' ?>" data-fmt="<?= h($k) ?>">
                    <input type="radio" name="format" value="<?= h($k) ?>" <?= $k === 'json' ? 'checked' : '' ?>>
                    <?= h($v['label']) ?>
                </label>
            <?php endforeach; ?>
        </div>
    </form>

    <div class="card">
        <h2>③ 转换选项</h2>
        <div class="opts">
            <label class="chk"><input type="checkbox" id="fill_merged" checked>拆分合并单元格（值复写到 x*y 每格）</label>
            <label class="chk"><input type="checkbox" id="extract_img" checked>提取图片到 demo/Img</label>
            <label class="chk opt-csv"><input type="checkbox" id="bom" checked>CSV/TSV 加 UTF-8 BOM</label>
            <label class="chk opt-json"><input type="checkbox" id="header">JSON 用第一行做键名</label>
            <label class="chk opt-json"><input type="checkbox" id="cells">JSON 输出完整单元格（含 img/公式）</label>
            <label class="chk">输出编码
                <select id="encoding" style="padding:4px 8px">
                    <option value="UTF-8">UTF-8</option>
                    <option value="GBK">GBK</option>
                    <option value="GB18030">GB18030</option>
                </select>
            </label>
        </div>
        <div class="note" id="hint"></div>
    </div>

    <div class="card">
        <div class="actions">
            <button type="button" id="run">转换</button>
            <button type="button" class="ghost" id="btnCopy">复制结果</button>
            <button type="button" class="ghost" id="btnDl">下载文件</button>
            <span id="spin" class="hide" style="color:var(--muted)">转换中…</span>
        </div>
        <div id="errBox" class="hide" style="margin-top:12px"></div>
    </div>

    <div class="card hide" id="result">
        <div class="bar">
            <div class="stats" id="stats"></div>
            <div class="tabs" id="tabs">
                <button data-view="text" class="on">源码</button>
                <button data-view="prev">预览</button>
            </div>
        </div>
        <pre id="pre"></pre>
        <iframe id="prev" class="hide" sandbox="allow-same-origin"></iframe>
        <div id="imgs" class="imgs"></div>
        <div class="note" id="dlNote"></div>
    </div>
    <?php endif; ?>
</div>

<div id="toast"></div>

<script>
(function () {
    const $ = (id) => document.getElementById(id);
    const state = { b64: null, text: '', name: 'export', mime: 'text/plain', format: 'json', view: 'text' };

    function toast(msg, bad) {
        const t = $('toast');
        t.textContent = msg;
        t.className = 'show' + (bad ? ' bad' : '');
        clearTimeout(t._t);
        t._t = setTimeout(() => { t.className = ''; }, 2200);
    }

    function curFmt() {
        const r = document.querySelector('#fmts input[name=format]:checked');
        return r ? r.value : 'json';
    }

    // 格式按钮高亮 + 选项联动
    function syncFmt() {
        const f = curFmt();
        document.querySelectorAll('#fmts .fmt').forEach(el => {
            el.classList.toggle('on', el.dataset.fmt === f);
        });
        document.querySelectorAll('.opt-csv').forEach(el => {
            el.classList.toggle('hide', !(f === 'csv' || f === 'tsv'));
        });
        document.querySelectorAll('.opt-json').forEach(el => {
            el.classList.toggle('hide', f !== 'json');
        });
        state.format = f;
        $('hint').textContent = {
            html: 'HTML 输出保留合并单元格（colspan / rowspan），并内联 <img> 引用 demo/Img 下的图片。',
            csv:  'CSV/TSV 会把合并区域的值复写到 x*y 的每一个格子，方便直接入库。',
            tsv:  'CSV/TSV 会把合并区域的值复写到 x*y 的每一个格子，方便直接入库。',
            json: 'JSON 默认二维数组；勾选「第一行做键名」后输出关联数组；勾选「完整单元格」输出 value/type/css/img/公式。',
        }[f] || '';
    }
    document.querySelectorAll('#fmts').forEach(el =>
        el.addEventListener('change', syncFmt));

    // 切换文件 → 拉取工作表列表
    async function loadSheets() {
        const fd = new FormData();
        fd.append('ajax', '1');
        fd.append('action', 'info');
        fd.append('file', $('file').value);
        try {
            const r = await fetch('demo.php', { method: 'POST', body: fd, headers: { 'X-Requested-With': 'XMLHttpRequest' } });
            const j = await r.json();
            const sel = $('sheet');
            sel.innerHTML = '';
            if (!j.ok || !j.sheets) {
                sel.innerHTML = '<option value="0">Sheet1</option>';
                return;
            }
            j.sheets.forEach(s => {
                const o = document.createElement('option');
                o.value = s.index;
                o.textContent = s.name + '  (' + s.rows + ' 行 × ' + s.cols + ' 列)';
                sel.appendChild(o);
            });
        } catch (e) {
            toast('读取工作表失败', true);
        }
    }

    // ---- 转换
    async function run() {
        $('run').disabled = true;
        $('spin').classList.remove('hide');
        $('errBox').classList.add('hide');

        const fd = new FormData();
        fd.append('ajax', '1');
        fd.append('action', 'convert');
        fd.append('file', $('file').value);
        fd.append('format', curFmt());
        fd.append('sheet', $('sheet').value);
        fd.append('fill_merged', $('fill_merged').checked ? '1' : '');
        fd.append('extract_img', $('extract_img').checked ? '1' : '');
        fd.append('bom', $('bom').checked ? '1' : '');
        fd.append('header', $('header').checked ? '1' : '0');
        fd.append('cells', $('cells').checked ? '1' : '');
        fd.append('encoding', $('encoding').value);

        try {
            const r = await fetch('demo.php', { method: 'POST', body: fd, headers: { 'X-Requested-With': 'XMLHttpRequest' } });
            const j = await r.json();
            if (!j.ok) throw new Error(j.error || '转换失败');

            state.b64   = j.b64;
            state.text  = j.text;                       // 复制用完整内容
            state.name  = j.file.replace(/\.xlsx$/i, '') + '.' + j.ext;
            state.mime  = j.mime;
            state.format = j.format;
            state.isUtf8 = j.isUtf8;

            // 统计徽章
            const m = j.meta;
            const s = [
                ['耗时', m.ms + ' ms'],
                ['尺寸', m.rows + ' 行 × ' + m.cols + ' 列'],
                ['工作表', m.sheet_name],
                ['合并区', m.merges],
                ['公式', m.formulas],
                ['图片', m.images],
                ['字节', m.bytes.toLocaleString()],
                ['编码', m.encoding],
            ];
            $('stats').innerHTML = s.map(([k, v]) =>
                '<span class="stat">' + k + ' <b>' + String(v) + '</b></span>').join('');

            // 源码视图
            $('pre').textContent = j.textCut !== null
                ? j.textCut + '\n\n…（内容过长，仅显示前 20000 字符；下载可获得完整内容）'
                : j.text;

            // 预览视图：仅 UTF-8 的完整 HTML 可直接渲染
            const canPrev = j.format === 'html' && j.isUtf8;
            document.querySelector('#tabs button[data-view=prev]').classList.toggle('hide', !canPrev);
            if (canPrev) {
                $('prev').srcdoc = j.text;
            } else if (j.format === 'html') {
                $('prev').srcdoc = '';
            }
            state.view = canPrev ? 'prev' : 'text';
            setView(state.view);

            // 图片清单
            $('imgs').innerHTML = (m.image_list || []).map(im =>
                '<figure><img src="' + im.path + '" alt=""><figcaption>' + (im.cell || '嵌入') + '</figcaption></figure>'
            ).join('');

            $('dlNote').textContent = j.tooBig
                ? '内容超过 5MB，请改用「下载文件」（会自动走服务端直出）。'
                : (j.isUtf8 ? '' : '注意：下载得到的是 ' + m.encoding + ' 原始字节；上方源码与复制内容是转回 UTF-8 后的显示版本。');

            $('result').classList.remove('hide');
            $('result').scrollIntoView({ behavior: 'smooth', block: 'start' });
            toast('转换完成');
        } catch (e) {
            const b = $('errBox');
            b.className = 'err';
            b.textContent = '出错了：' + e.message;
            b.classList.remove('hide');
            toast('转换失败', true);
        } finally {
            $('run').disabled = false;
            $('spin').classList.add('hide');
        }
    }

    function setView(v) {
        state.view = v;
        document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.view === v));
        $('pre').classList.toggle('hide', v !== 'text');
        $('prev').classList.toggle('hide', v !== 'prev');
    }

    // ---- 复制
    async function copy() {
        if (!state.text) { toast('还没有可复制的内容', true); return; }
        try {
            await navigator.clipboard.writeText(state.text);
            toast('已复制到剪贴板');
        } catch (e) {
            // 降级方案：textarea + execCommand（http 环境或非安全上下文下 clipboard 不可用）
            const ta = document.createElement('textarea');
            ta.value = state.text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            const ok = document.execCommand('copy');
            document.body.removeChild(ta);
            toast(ok ? '已复制到剪贴板' : '复制失败，请手动选择', !ok);
        }
    }

    // ---- 下载（用 base64 还原原始字节，保证 GBK 也不损坏）
    function download() {
        if (!state.b64) { toast('内容过大，请用服务端下载', true); return; }
        const bin = atob(state.b64);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        const url = URL.createObjectURL(new Blob([arr], { type: state.mime }));
        const a = document.createElement('a');
        a.href = url;
        a.download = state.name;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 3000);
        toast('已下载 ' + state.name);
    }

    $('run').addEventListener('click', run);
    $('btnCopy').addEventListener('click', copy);
    $('btnDl').addEventListener('click', download);
    $('file').addEventListener('change', loadSheets);
    document.querySelectorAll('#tabs button').forEach(b =>
        b.addEventListener('click', () => setView(b.dataset.view)));

    syncFmt();
    loadSheets();
})();
</script>
</body>
</html>
