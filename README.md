# SimpleXLSX Pro —— 中文增强套件

在 [`shuchkin/simplexlsx`](https://github.com/shuchkin/simplexlsx) 之上补齐三块能力：**公式计算**、**图片提取**、**合并单元格与多格式导出**，并修复原库 7 处实测确认的缺陷。

纯 PHP、零依赖，PHP 7.1+ / 8.x 均可运行。

---

## 1. 目录结构

```
src/
  SimpleXLSX.php        # 原库主类（已打补丁）
  SimpleXLSXEx.php      # 原库样式扩展（已打补丁）
  SimpleXLSXPro.php     # ★ 增强子类：公式 / 图片 / 合并 / 导出
  XLSXFormula.php       # ★ 公式计算引擎
examples/
  00-smoke_test.php           # 冒烟测试（24 项断言，全通过）
  01-quick_start.php          # 快速上手
  02-images.php               # 图片提取（含 WPS 内嵌图）
  03-formulas.php             # 公式计算
  04-export_csv_tsv_json.php  # CSV / TSV / JSON 导出
  05-export_html_merged.php   # HTML 导出（保留合并）
  06-cli_convert.php          # 命令行批量转换工具
demo/
  demo.xlsx                   # 测试文件：合并单元格 + 公式 + 浮动图片
  demo_wps.xlsx               # 测试文件：WPS 单元格内嵌图片（DISPIMG）
simplexlsx-fixes.patch        # 原库补丁（unified diff）
```

## 2. 安装

```bash
# 方式 A：直接用本套件（src/ 里已是打过补丁的完整源码）
cp -r src/* your-project/simplexlsx/

# 方式 B：你已经在用 composer 装的 simplexlsx，只打补丁
cd vendor/shuchkin/simplexlsx
patch -p1 < /path/to/simplexlsx-fixes.patch
# 然后把 SimpleXLSXPro.php + XLSXFormula.php 复制到你的类库目录
```

代码里引入：

```php
require 'src/SimpleXLSXPro.php';   // 会自动加载 SimpleXLSX.php 与 XLSXFormula.php
use Shuchkin\SimpleXLSXPro;
```

## 3. 30 秒上手

```php
use Shuchkin\SimpleXLSXPro;

$xlsx = SimpleXLSXPro::parse('demo.xlsx');
if (!$xlsx) { exit(SimpleXLSXPro::parseError()); }

$xlsx->extractImages('Img');                      // 图片 → Img/ 目录
echo $xlsx->calculateCell('D6');                  // 745（公式算出来的）
echo $xlsx->toCSV(0, true);                       // CSV，合并单元格已拆分复写
echo $xlsx->toJSON(0, ['cells' => true]);         // JSON，每格带 img / merge / formula
echo $xlsx->toHTMLMerged(0, ['images' => true]);  // HTML，真正的 colspan/rowspan
```

命令行：

```bash
cd demo && php ../examples/06-cli_convert.php demo.xlsx --format=csv
```

## 4. 修复的 7 个缺陷

| # | 位置 | 现象 |
| --- | --- | --- |
| 1 | `SimpleXLSXEx::readRowsEx()` | `$limit` 完全失效（`break` 应为 `return`），想取前 100 行却拿到整张表 |
| 2 | `SimpleXLSXEx::valueEx()` | CSS 引用赋值导致字符串每次调用翻倍累加 |
| 3 | `SimpleXLSX::toHTML()/toHTMLEx()` | README 写的 `$limit` 参数源码里根本不存在 |
| 4 | `SimpleXLSX::parseError()` | 静态错误状态永不重置：失败一次后，后续成功仍报旧错 |
| 5 | `SimpleXLSXEx::valueEx()` | 样式索引越界时 PHP 8 Warning |
| 6 | `SimpleXLSXEx::readRowsEx()` | 单元格列数超出 `dimension` 时 `$cols` 未同步扩展 → Warning 洪水 |
| 7 | `SimpleXLSXEx::readRowsEx()` | 跳列写入导致列顺序错乱（…F3 H3 G3） |

## 5. 新增能力速查

### 公式

```php
$xlsx->calculateCell('D6');                  // 算单元格
$xlsx->calculate('SUM(A1:A5)*2');            // 算任意公式
$xlsx->calculatedCells(0);                   // 全部公式单元格
$xlsx->calculateFormulas = false;            // 关掉计算（最快）
```

60+ 函数：数学 / 统计 / 逻辑 / 文本 / 查找 / 日期，支持跨表引用（含中文表名）、区域引用、错误值传播、`#CIRC!` 循环引用检测。

### 图片

```php
$map = $xlsx->extractImages('Img');          // [sheet][row][col] => 'Img/xxx.png'
$xlsx->imageList();                          // 清单
$xlsx->imageAt(0, 5, 2);                     // 按坐标取
```

支持浮动图片（drawing 锚点）与 WPS 内嵌图片（`=DISPIMG("ID_xxx",1)`）。`rowsPro()` / `toJSON(['cells'=>true])` 的每个单元格都带 `img` 字段。

### 合并单元格

```php
$xlsx->mergeCells(0);        // [['ref'=>'A1:E1','x1'=>..,'y1'=>..,'w'=>5,'h'=>1], ...]
$xlsx->mergeAt(0, 4, 3);     // 单格合并信息
$xlsx->fillMerged($rows);    // 给已有二维数组做拆分复写
```

- CSV / TSV / JSON：值拆分复写到 x\*y 每个格
- HTML：输出 `colspan` / `rowspan`，被合并的格不重复输出

### 导出

| 方法 | 说明 |
| --- | --- |
| `toArray($sheet, $fillMerged, $limit)` | 二维数组 |
| `toCSV($sheet, $fillMerged, $bom)` / `saveCSV()` | CSV（默认带 UTF-8 BOM） |
| `toTSV()` / `saveTSV()` | TSV |
| `toJSON($sheet, $options)` / `saveJSON()` | JSON，选项见下 |
| `toHTMLMerged($sheet, $options)` | 保留合并 + 图片的 HTML |
| `SimpleXLSXPro::convert($file, $format, $opt)` | 一步到位 |

`toJSON` 选项：`fill_merged` `header`（可指定行号）`cells` `meta` `pretty` `flags`。

## 6. 运行测试

```bash
cd demo
php ../examples/00-smoke_test.php     # 24 项断言，应全部 [OK]
php ../examples/01-quick_start.php
php ../examples/02-images.php demo_wps.xlsx
```

## 7. 许可

补丁与扩展代码遵循原库许可（MIT）。
