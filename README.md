# SimpleXLSX Pro —— 中文增强套件

在 [`shuchkin/simplexlsx`](https://github.com/shuchkin/simplexlsx) 之上补齐三块能力：**公式计算**、**图片提取**、**合并单元格与多格式导出**，并修复原库 7 处实测确认的缺陷。

纯 PHP、零依赖，PHP 7.1+ / 8.x 均可运行。更多说明书见html文件，所有代码,说明书,测试用例均使用AI(仅供参考)。

# SimpleXLSX / SimpleXLSXPro

> 纯 PHP 的 Excel xlsx 读取与多格式导出套件 · Pure-PHP xlsx reader & exporter suite · Набор PHP-библиотек для чтения и экспорта xlsx

[中文](#中文) | [English](#english) | [Русский](#русский)

---

<a id="中文"></a>
## 中文

**SimpleXLSX** 是一套纯 PHP 的 Excel 2007+ `.xlsx` 解析库，无需框架、无需数据库，单文件引入即可使用，兼容 PHP 7.0+。

### 组成

| 文件 | 说明 |
|---|---|
| `src/SimpleXLSX.php` | 基础版：读取单元格数据、多工作表、日期 |
| `src/SimpleXLSXEx.php` | 样式增强版：字体/填充/对齐/边框转 CSS |
| `src/SimpleXLSXPro.php` | 专业版：合并单元格、图片提取、多格式导出 |
| `src/XLSXFormula.php` | 公式引擎：60+ 函数、跨表引用实时计算 |
| `demo.php` | 单文件在线演示台（AJAX 转换 + 预览 + 下载） |
| `demo/` | 7 个功能样例 xlsx（合并、公式、多表、图片等） |

### 特性

- 解析 xlsx 为 PHP 二维数组；支持多工作表、日期时间、富文本、超链接
- 合并单元格：CSV/TSV/JSON 值复写到每格，HTML 保留 rowspan/colspan
- 内置公式引擎，无缓存值也能实时计算（SUM/IF/VLOOKUP/跨表引用等）
- 提取 xlsx 内嵌图片（含 WPS DISPIMG 内嵌图）到指定目录
- 导出 CSV / TSV / JSON / HTML，支持 UTF-8/GBK/GB18030 编码

### 快速上手

```php
require_once __DIR__ . '/src/SimpleXLSXPro.php';
$xlsx = Shuchkin\SimpleXLSXPro::parse('book.xlsx');
foreach ($xlsx->rows() as $row) { print_r($row); }
echo $xlsx->toHTMLMerged();   // HTML（保留合并 + 图片）
echo $xlsx->toCSV();          // CSV / toTSV() / toJSON()
```

### 演示台

把整个目录放到 PHP 站点（保证 `demo/` 可写），访问 `demo.php`：选择文件与格式 → AJAX 转换 → 复制 / 下载。详见 `docs/` 下三语开发文档。

---

<a id="english"></a>
## English

**SimpleXLSX** is a pure-PHP parser for Excel 2007+ `.xlsx` workbooks — no framework, no database, drop-in single files, PHP 7.0+ compatible.

### Layout

| File | Purpose |
|---|---|
| `src/SimpleXLSX.php` | Core: cell data, multiple sheets, dates |
| `src/SimpleXLSXEx.php` | Styles: font/fill/alignment/border → CSS |
| `src/SimpleXLSXPro.php` | Pro: merged cells, image extraction, exporters |
| `src/XLSXFormula.php` | Formula engine: 60+ functions, cross-sheet evaluation |
| `demo.php` | Single-file online demo (AJAX convert + preview + download) |
| `demo/` | 7 sample xlsx files (merges, formulas, sheets, images…) |

### Features

- Parse xlsx into PHP arrays; multiple sheets, dates, rich text, hyperlinks
- Merged cells: values copied to every cell in CSV/TSV/JSON; real rowspan/colspan in HTML
- Built-in formula engine evaluates SUM/IF/VLOOKUP/cross-sheet refs even without cached values
- Extract embedded images (incl. WPS DISPIMG) to a directory
- Export to CSV / TSV / JSON / HTML with UTF-8/GBK/GB18030 encodings

### Quick start

```php
require_once __DIR__ . '/src/SimpleXLSXPro.php';
$xlsx = Shuchkin\SimpleXLSXPro::parse('book.xlsx');
foreach ($xlsx->rows() as $row) { print_r($row); }
echo $xlsx->toHTMLMerged();   // HTML (merges + images)
echo $xlsx->toCSV();          // CSV / toTSV() / toJSON()
```

### Demo

Deploy the folder on any PHP host (make `demo/` writable) and open `demo.php`: pick a file and format → AJAX convert → copy / download. Full manuals in `docs/` (zh / en / ru).

---

<a id="русский"></a>
## Русский

**SimpleXLSX** — чистый PHP-парсер книг Excel 2007+ `.xlsx`: без фреймворков и базы данных, подключается одним файлом, совместим с PHP 7.0+.

### Состав

| Файл | Назначение |
|---|---|
| `src/SimpleXLSX.php` | Базовая версия: данные ячеек, листы, даты |
| `src/SimpleXLSXEx.php` | Стили: шрифт/заливка/выравнивание/границы → CSS |
| `src/SimpleXLSXPro.php` | Pro: объединённые ячейки, извлечение изображений, экспорт |
| `src/XLSXFormula.php` | Движок формул: 60+ функций, межлистовые ссылки |
| `demo.php` | Онлайн-демо одним файлом (AJAX-конвертация + предпросмотр) |
| `demo/` | 7 примеров xlsx (объединения, формулы, листы, картинки…) |

### Возможности

- Разбор xlsx в PHP-массивы; листы, даты, форматированный текст, гиперссылки
- Объединённые ячейки: значение копируется в каждую ячейку (CSV/TSV/JSON); настоящие rowspan/colspan в HTML
- Встроенный движок формул: SUM/IF/VLOOKUP/межлистовые ссылки вычисляются даже без кэша
- Извлечение встроенных изображений (включая WPS DISPIMG) в каталог
- Экспорт в CSV / TSV / JSON / HTML, кодировки UTF-8/GBK/GB18030

### Быстрый старт

```php
require_once __DIR__ . '/src/SimpleXLSXPro.php';
$xlsx = Shuchkin\SimpleXLSXPro::parse('book.xlsx');
foreach ($xlsx->rows() as $row) { print_r($row); }
echo $xlsx->toHTMLMerged();   // HTML (объединения + картинки)
echo $xlsx->toCSV();          // CSV / toTSV() / toJSON()
```

### Демо

Разместите папку на любом PHP-хостинге (каталог `demo/` должен быть доступен для записи) и откройте `demo.php`: выбор файла и формата → AJAX-конвертация → копирование / скачивание. Полные руководства — в `docs/` (zh / en / ru).

---

## License / 许可 / Лицензия

MIT — based on [shuchkin/simplexlsx](https://github.com/shuchkin/simplexlsx)
