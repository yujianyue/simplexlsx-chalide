<?php
/**
* SimpleXLSXPro - SimpleXLSX 的功能增强子类
*
* 在父类之上新增：
*   1. 图片提取：把 xlsx 内嵌图片导出到 Img 目录，并在单元格里返回 img 路径
*      （同时支持浮动图片 drawing 与 WPS 单元格内嵌图片 DISPIMG）
*   2. 公式计算：内置 XLSXFormula 引擎，支持 60+ 函数、跨表引用、循环引用检测
*   3. 合并单元格：解析 mergeCells，CSV/TSV/JSON 把值拆分复写到 x*y 每个格，
*      HTML 输出保留真正的 rowspan/colspan
*   4. 多种导出：CSV / TSV / JSON / HTML(带合并与图片)
*
* @package Shuchkin
* @author  SimpleXLSX Pro extension
*/
namespace Shuchkin;
if (!class_exists('Shuchkin\\SimpleXLSX', false)) {
foreach ([
__DIR__ . '/SimpleXLSX.php',
dirname(__DIR__) . '/SimpleXLSX.php',
dirname(__DIR__, 2) . '/simplexlsx/src/SimpleXLSX.php',
] as $candidate) {
if (is_file($candidate)) {
require_once $candidate;
break;
}
}
}
require_once __DIR__ . '/XLSXFormula.php';
class SimpleXLSXPro extends SimpleXLSX
{
/** @var string 图片输出目录（相对当前工作目录） */
public $imageDir = 'Img';
/** @var bool 是否让计算值覆盖空的缓存值 */
public $preferCalculated = true;
/** @var bool 是否在 readRowsPro() 中计算公式 */
public $calculateFormulas = true;
/** @var array [sheet][row][col] => 图片相对路径 */
protected $imageMap = [];
/** @var array 图片清单 */
protected $imageList = [];
/** @var array WPS DISPIMG 图片 ID => 路径 */
protected $dispImages = [];
/** @var array [sheet] => [['ref','x1','y1','x2','y2','w','h'], ...] */
protected $mergeList = [];
/** @var array [sheet][row][col] => ['w','h','ref','master'] */
protected $mergeMap = [];
/** @var array [sheet]['A1'] => ['value','formula','type'] */
protected $cellMap = [];
/** @var XLSXFormula|null */
protected $formulaEngine;
/** @var bool 图片是否已提取 */
protected $imagesExtracted = false;
// ------------------------------------------------------------------ 入口
/**
* @param string $filename
* @param bool   $is_data
* @param bool   $debug
*
* @return self|false
*/
public static function parse($filename, $is_data = false, $debug = false)
{
$xlsx = new self($filename, $is_data, $debug);
if ($xlsx->success()) {
self::parseError(false, true);
self::parseErrno(false, true);
return $xlsx;
}
self::parseError($xlsx->error());
self::parseErrno($xlsx->errno());
return false;
}
/**
* @param string $data 二进制内容
*
* @return self|false
*/
public static function parseData($data, $debug = false)
{
return self::parse($data, true, $debug);
}
// ------------------------------------------------------------------ 图片
/**
* 提取全部内嵌图片到目录，并返回 [sheet][row][col] => 路径 的映射
*
* @param string|null $dir    输出目录，默认 $this->imageDir（相对 CWD）
* @param bool        $returnRelative 记录相对路径还是绝对路径
*
* @return array
*/
public function extractImages($dir = null, $returnRelative = true)
{
$dir = $dir === null ? $this->imageDir : $dir;
if ($dir === '' || $dir === null) {
return $this->imageMap;
}
if (!is_dir($dir) && !@mkdir($dir, 0777, true) && !is_dir($dir)) {
$this->error(8, 'Cannot create image directory ' . $dir);
return $this->imageMap;
}
$base = dirname($this->sheetFiles[0] ?? 'xl/worksheets/sheet1.xml');
foreach ($this->sheetRels as $index => $xml) {
foreach ($xml->Relationship as $rel) {
$relType = basename(trim((string)$rel['Type']));
$target = self::getTarget(
dirname(isset($this->sheetFiles[$index]) ? $this->sheetFiles[$index] : $base),
(string)$rel['Target']
);
if ($relType === 'drawing') {
$this->parseDrawing($index, $target, $dir, $returnRelative);
} elseif ($relType === 'cellimages' || $relType === 'cellImage') {
$this->parseCellImages($target, $dir, $returnRelative);
}
}
}
$this->imagesExtracted = true;
return $this->imageMap;
}
/**
* 解析 drawingN.xml，按锚点定位到单元格
*/
protected function parseDrawing($sheetIndex, $drawingPath, $dir, $returnRelative)
{
if (!$this->entryExists($drawingPath)) {
return;
}
$xml = $this->getEntryXML($drawingPath); // 会释放该 entry
if (!$xml) {
return;
}
// drawing 的关系文件：rId => media 路径
$relsPath = dirname($drawingPath) . '/_rels/' . basename($drawingPath) . '.rels';
$rels = [];
if ($this->entryExists($relsPath)) {
$rx = $this->getEntryXML($relsPath);
if ($rx) {
foreach ($rx->Relationship as $r) {
$rels[(string)$r['Id']] = self::getTarget(dirname($drawingPath), (string)$r['Target']);
}
}
}
$seq = 0;
foreach ($xml->children() as $anchor) {
$tag = $anchor->getName();
if (stripos($tag, 'Anchor') === false) {
continue;
}
$col = isset($anchor->from->col) ? (int)$anchor->from->col : 0;
$row = isset($anchor->from->row) ? (int)$anchor->from->row : 0;
$blips = $anchor->xpath('.//blip');
if (!$blips) {
$blips = $anchor->xpath('.//*/blip');
}
if (!$blips) {
continue;
}
foreach ($blips as $blip) {
$rid = (string)$blip['embed'];
if ($rid === '') {
$rid = (string)$blip['link']; // 外链图片，多数情况取不到内容
}
if ($rid === '' || !isset($rels[$rid])) {
continue;
}
$media = $rels[$rid];
$path = $this->saveMedia($media, $dir, 'sheet' . ($sheetIndex + 1)
. '_' . XLSXFormula::indexToCol($col) . ($row + 1) . '_' . (++$seq), $returnRelative);
if ($path === null) {
continue;
}
$this->imageMap[$sheetIndex][$row][$col] = $path;
$this->imageList[] = [
'sheet'    => $sheetIndex,
'sheet_name' => $this->sheetName($sheetIndex),
'col'      => $col,
'row'      => $row,
'cell'     => XLSXFormula::indexToCol($col) . ($row + 1),
'path'     => $path,
'source'   => $media,
'name'     => (string)($blip['descr'] ?? ''),
];
}
}
}
/**
* 解析 WPS 单元格内嵌图片 xl/cellimages.xml（配合 =DISPIMG("ID_xxx",1)）
*/
protected function parseCellImages($path, $dir, $returnRelative)
{
if (!$this->entryExists($path)) {
return;
}
$xml = $this->getEntryXML($path);
if (!$xml) {
return;
}
$relsPath = dirname($path) . '/_rels/' . basename($path) . '.rels';
$rels = [];
if ($this->entryExists($relsPath)) {
$rx = $this->getEntryXML($relsPath);
if ($rx) {
foreach ($rx->Relationship as $r) {
$rels[(string)$r['Id']] = self::getTarget(dirname($path), (string)$r['Target']);
}
}
}
$seq = 0;
foreach ($xml->xpath('.//pic') as $pic) {
$name = '';
$nv = $pic->xpath('.//cNvPr');
if ($nv) {
$name = (string)$nv[0]['name'];
}
$blips = $pic->xpath('.//blip');
if (!$blips) {
continue;
}
$rid = (string)$blips[0]['embed'];
if ($rid === '' || !isset($rels[$rid])) {
continue;
}
$save = $this->saveMedia($rels[$rid], $dir, 'embed_' . (++$seq), $returnRelative);
if ($save === null) {
continue;
}
if ($name !== '') {
$this->dispImages[$name] = $save;
}
$this->imageList[] = [
'sheet'      => null,
'sheet_name' => '',
'col'        => null,
'row'        => null,
'cell'       => '',
'path'       => $save,
'source'     => $rels[$rid],
'name'       => $name,
];
}
}
/**
* 把一个 media entry 写成文件
*
* @return string|null 成功返回路径（相对或绝对），失败返回 null
*/
/** @var array media 内部路径 => 已保存的文件路径（避免同一张图重复落盘） */
protected $mediaCache = [];
/**
* 把一个 media entry 写成文件
*
* @return string|null 成功返回路径（相对或绝对），失败返回 null
*/
protected function saveMedia($mediaPath, $dir, $prefix, $returnRelative)
{
$cacheKey = $mediaPath . '|' . ($returnRelative ? 'r' : 'a');
if (isset($this->mediaCache[$cacheKey])) {
return $this->mediaCache[$cacheKey]; // 同一张图片被多处引用
}
if (!$this->entryExists($mediaPath)) {
return null;
}
$data = $this->getEntryData($mediaPath);
if ($data === false || $data === '' || $data === null) {
return null;
}
$ext = strtolower(pathinfo($mediaPath, PATHINFO_EXTENSION));
if ($ext === '') {
$ext = 'bin';
}
$file = $prefix . '.' . $ext;
$full = rtrim($dir, '/\\') . DIRECTORY_SEPARATOR . $file;
if (@file_put_contents($full, $data) === false) {
return null;
}
// 释放该 entry，避免大图常驻内存（路径已缓存，后续引用直接复用）
$this->deleteEntry($mediaPath);
$result = $returnRelative ? str_replace('\\', '/', $dir . '/' . $file) : $full;
$this->mediaCache[$cacheKey] = $result;
return $result;
}
/**
* 是否已提取图片
*/
public function isImagesExtracted()
{
return $this->imagesExtracted;
}
/**
* 图片清单
*/
public function imageList()
{
return $this->imageList;
}
/**
* [sheet][row][col] => 路径
*/
public function imageMap()
{
return $this->imageMap;
}
/**
* 取某单元格的图片路径
*/
public function imageAt($sheetIndex, $col, $row)
{
return isset($this->imageMap[$sheetIndex][$row][$col]) ? $this->imageMap[$sheetIndex][$row][$col] : '';
}
// ------------------------------------------------------------------ 合并单元格
/**
* 解析合并区域
*
* @return array [['ref'=>'A1:B2','x1'=>0,'y1'=>0,'x2'=>1,'y2'=>1,'w'=>2,'h'=>2], ...]
*/
public function mergeCells($worksheetIndex = 0)
{
if (isset($this->mergeList[$worksheetIndex])) {
return $this->mergeList[$worksheetIndex];
}
$this->mergeList[$worksheetIndex] = [];
$this->mergeMap[$worksheetIndex] = [];
$ws = $this->worksheet($worksheetIndex);
if ($ws === false || !isset($ws->mergeCells)) {
return [];
}
foreach ($ws->mergeCells->mergeCell as $mc) {
$ref = (string)$mc['ref'];
if (strpos($ref, ':') === false) {
continue;
}
list($from, $to) = explode(':', $ref);
$p1 = $this->getIndex($from);
$p2 = $this->getIndex($to);
if ($p1[0] < 0 || $p2[0] < 0) {
continue;
}
$x1 = min($p1[0], $p2[0]);
$x2 = max($p1[0], $p2[0]);
$y1 = min($p1[1], $p2[1]);
$y2 = max($p1[1], $p2[1]);
$item = [
'ref' => $ref,
'x1'  => $x1,
'y1'  => $y1,
'x2'  => $x2,
'y2'  => $y2,
'w'   => $x2 - $x1 + 1,
'h'   => $y2 - $y1 + 1,
];
$this->mergeList[$worksheetIndex][] = $item;
for ($y = $y1; $y <= $y2; $y++) {
for ($x = $x1; $x <= $x2; $x++) {
$this->mergeMap[$worksheetIndex][$y][$x] = [
'ref'    => $ref,
'w'      => $item['w'],
'h'      => $item['h'],
'master' => ($x === $x1 && $y === $y1),
];
}
}
}
return $this->mergeList[$worksheetIndex];
}
/**
* 某单元格的合并信息，未合并返回 null
*/
public function mergeAt($sheetIndex, $col, $row)
{
$this->mergeCells($sheetIndex);
return isset($this->mergeMap[$sheetIndex][$row][$col]) ? $this->mergeMap[$sheetIndex][$row][$col] : null;
}
// ------------------------------------------------------------------ 公式
/**
* 构建单元格原始值 / 公式索引（懒加载，公式计算必需）
*/
protected function buildCellMap($worksheetIndex)
{
if (isset($this->cellMap[$worksheetIndex])) {
return;
}
$this->cellMap[$worksheetIndex] = [];
$ws = $this->worksheet($worksheetIndex);
if ($ws === false) {
return;
}
foreach ($ws->sheetData->row as $row) {
foreach ($row->c as $c) {
$name = strtoupper((string)$c['r']);
if ($name === '') {
continue;
}
$this->cellMap[$worksheetIndex][$name] = [
'value'   => $this->value($c),
'formula' => (string)$c->f,
'type'    => (string)$c['t'],
];
}
}
}
/**
* 公式引擎（懒加载）
*/
public function formula()
{
if ($this->formulaEngine === null) {
$self = $this;
$this->formulaEngine = new XLSXFormula(
function ($sheetIndex, $coord) use ($self) {
return $self->rawCell($sheetIndex, $coord);
},
0
);
$this->formulaEngine->setSheetNames($this->sheetNames());
}
return $this->formulaEngine;
}
/**
* 供公式引擎取单元格原始数据（public 以便闭包调用）
*/
public function rawCell($sheetIndex, $coord)
{
$this->buildCellMap($sheetIndex);
return isset($this->cellMap[$sheetIndex][$coord]) ? $this->cellMap[$sheetIndex][$coord] : null;
}
/**
* 计算某个单元格的公式结果
*
* @param string   $coord  A1
* @param int|null $sheetIndex
*
* @return mixed 出错时返回 '#DIV/0!' 之类的错误码字符串
*/
public function calculateCell($coord, $sheetIndex = null)
{
$sheetIndex = $sheetIndex === null ? 0 : (int)$sheetIndex;
$v = $this->formula()->cellValue(strtoupper($coord), $sheetIndex);
return $v instanceof XLSXFormulaError ? $v->code : $v;
}
/**
* 直接计算一段公式文本
*/
public function calculate($formula, $sheetIndex = 0)
{
$v = $this->formula()->calculate($formula, $sheetIndex);
return $v instanceof XLSXFormulaError ? $v->code : $v;
}
/**
* 全部含公式的单元格 => 计算值
*/
public function calculatedCells($worksheetIndex = 0)
{
$this->buildCellMap($worksheetIndex);
$out = [];
foreach ($this->cellMap[$worksheetIndex] as $coord => $info) {
if ($info['formula'] !== '') {
$out[$coord] = [
'formula' => $info['formula'],
'cached'  => $info['value'],
'value'   => $this->calculateCell($coord, $worksheetIndex),
];
}
}
return $out;
}
// ------------------------------------------------------------------ 增强读取
/**
* rowsEx() 的增强版：每个单元格额外返回
*   img        图片路径（'Img/xxx.png' 或 ''）
*   merge      合并信息 ['ref','w','h','master'] 或 null
*   calculated 公式计算值（无公式时为 null）
*
* @return \Generator
*/
public function readRowsPro($worksheetIndex = 0, $limit = 0)
{
$this->mergeCells($worksheetIndex);
$y = 0;
foreach ($this->readRowsEx($worksheetIndex, $limit) as $row) {
$x = 0;
foreach ($row as &$c) {
$c['img'] = isset($this->imageMap[$worksheetIndex][$y][$x])
? $this->imageMap[$worksheetIndex][$y][$x] : '';
// WPS 内嵌图片 =DISPIMG("ID_xxx",1)
if ($c['img'] === '' && $c['f'] !== '' && preg_match('/DISPIMG\(\s*"([^"]+)"/i', $c['f'], $m)) {
$c['img'] = isset($this->dispImages[$m[1]]) ? $this->dispImages[$m[1]] : '';
}
$c['merge'] = isset($this->mergeMap[$worksheetIndex][$y][$x])
? $this->mergeMap[$worksheetIndex][$y][$x] : null;
$c['calculated'] = null;
if ($c['f'] !== '' && $this->calculateFormulas) {
$v = $this->formula()->cellValue(strtoupper($c['name']), $worksheetIndex);
$c['calculated'] = $v instanceof XLSXFormulaError ? $v->code : $v;
if ($this->preferCalculated && ($c['value'] === '' || $c['value'] === null)
&& $c['calculated'] !== null) {
$c['value'] = $c['calculated'];
}
}
$x++;
}
unset($c);
yield $row;
$y++;
}
}
/**
* @return array
*/
public function rowsPro($worksheetIndex = 0, $limit = 0)
{
return iterator_to_array($this->readRowsPro($worksheetIndex, $limit), false);
}
// ------------------------------------------------------------------ 导出：数组
/**
* 取二维数组，$fillMerged = true 时把合并单元格的值复写到区域内每一个格
*
* @param int  $worksheetIndex
* @param bool $fillMerged
* @param int  $limit
*
* @return array
*/
public function toArray($worksheetIndex = 0, $fillMerged = true, $limit = 0)
{
$rows = $this->rows($worksheetIndex, $limit);
// 公式单元格：无缓存值时用计算结果补上
if ($this->calculateFormulas) {
$this->buildCellMap($worksheetIndex);
if (!empty($this->cellMap[$worksheetIndex])) {
foreach ($rows as $y => $row) {
foreach ($row as $x => $v) {
if ($v !== '' && $v !== null) {
continue;
}
$coord = strtoupper(XLSXFormula::indexToCol($x) . ($y + 1));
if (!empty($this->cellMap[$worksheetIndex][$coord]['formula'])) {
$rows[$y][$x] = $this->calculateCell($coord, $worksheetIndex);
}
}
}
}
}
if ($fillMerged) {
$rows = $this->fillMerged($rows, $worksheetIndex);
}
return $this->normalize($rows);
}
/**
* 把合并区域的值拆分复写到 x*y 的每一个格
*/
public function fillMerged(array $rows, $worksheetIndex = 0)
{
foreach ($this->mergeCells($worksheetIndex) as $m) {
$value = isset($rows[$m['y1']][$m['x1']]) ? $rows[$m['y1']][$m['x1']] : '';
for ($y = $m['y1']; $y <= $m['y2']; $y++) {
for ($x = $m['x1']; $x <= $m['x2']; $x++) {
if (!isset($rows[$y])) {
$rows[$y] = [];
}
$rows[$y][$x] = $value;
}
}
}
return $rows;
}
/**
* 补齐每行长度，保证矩形矩阵
*/
protected function normalize(array $rows)
{
$width = 0;
foreach ($rows as $row) {
$width = max($width, count($row));
}
foreach ($rows as $y => $row) {
if (count($row) < $width) {
for ($x = count($row); $x < $width; $x++) {
$rows[$y][$x] = '';
}
}
ksort($rows[$y]);
}
return $rows;
}
// ------------------------------------------------------------------ 导出：CSV / TSV
/**
* @param int  $worksheetIndex
* @param bool $fillMerged
* @param bool $bom 是否加 UTF-8 BOM（Excel 直接双击打开不乱码）
*
* @return string
*/
public function toCSV($worksheetIndex = 0, $fillMerged = true, $bom = true)
{
return $this->toDelimited(',', $worksheetIndex, $fillMerged, $bom);
}
public function toTSV($worksheetIndex = 0, $fillMerged = true, $bom = false)
{
return $this->toDelimited("\t", $worksheetIndex, $fillMerged, $bom);
}
protected function toDelimited($delimiter, $worksheetIndex, $fillMerged, $bom)
{
$rows = $this->toArray($worksheetIndex, $fillMerged);
$fh = fopen('php://memory', 'w+b');
if ($bom) {
fwrite($fh, "\xEF\xBB\xBF");
}
foreach ($rows as $row) {
fputcsv($fh, $row, $delimiter);
}
rewind($fh);
$out = stream_get_contents($fh);
fclose($fh);
return $out;
}
public function saveCSV($path, $worksheetIndex = 0, $fillMerged = true)
{
return file_put_contents($path, $this->toCSV($worksheetIndex, $fillMerged)) !== false;
}
public function saveTSV($path, $worksheetIndex = 0, $fillMerged = true)
{
return file_put_contents($path, $this->toTSV($worksheetIndex, $fillMerged)) !== false;
}
// ------------------------------------------------------------------ 导出：JSON
/**
* JSON 导出
*
* 选项：
*   fill_merged  true   合并单元格值复写到 x*y 每格
*   header       false  用第一行作为键，输出对象数组
*   cells        false  输出完整单元格信息（value/img/merge/formula/calculated）
*   meta         false  附带 sheet / dimension / merges / images 元信息
*   pretty       true   格式化输出
*   flags        int    json_encode 附加标志
*
* @param int   $worksheetIndex
* @param array $options
*
* @return string
*/
public function toJSON($worksheetIndex = 0, array $options = [])
{
$opt = array_merge([
'fill_merged' => true,
'header'      => false,   // true = 第一行做键；数字 = 指定行号做键
'cells'       => false,
'meta'        => false,
'pretty'      => true,
'flags'       => 0,
], $options);
$flags = JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | (int)$opt['flags'];
if ($opt['pretty']) {
$flags |= JSON_PRETTY_PRINT;
}
if ($opt['cells']) {
$data = $this->toCellsArray($worksheetIndex, (bool)$opt['fill_merged']);
} else {
$data = $this->toArray($worksheetIndex, (bool)$opt['fill_merged']);
if ($opt['header'] !== false && count($data) > 1) {
// header 可以是 true（第一行）或行号（从 0 开始）
$headerRow = is_numeric($opt['header']) ? (int)$opt['header'] : 0;
$header = isset($data[$headerRow]) ? $data[$headerRow] : array_shift($data);
$data = array_slice($data, $headerRow + 1);
$assoc = [];
foreach ($data as $row) {
$item = [];
foreach ($header as $i => $key) {
$k = ($key === '' || $key === null) ? 'col' . ($i + 1) : (string)$key;
$item[$k] = isset($row[$i]) ? $row[$i] : '';
}
$assoc[] = $item;
}
$data = $assoc;
}
}
if ($opt['meta']) {
$data = [
'sheet'     => $this->sheetName($worksheetIndex),
'index'     => $worksheetIndex,
'dimension' => $this->dimension($worksheetIndex),
'merges'    => $this->mergeCells($worksheetIndex),
'images'    => $this->imageList(),
'data'      => $data,
];
}
return json_encode($data, $flags);
}
/**
* 完整单元格信息（用于需要图片路径 / 合并信息的场景）
*/
public function toCellsArray($worksheetIndex = 0, $fillMerged = true)
{
$out = [];
$mergeValues = [];
if ($fillMerged) {
foreach ($this->mergeCells($worksheetIndex) as $m) {
$mergeValues[$m['ref']] = null;
}
}
foreach ($this->readRowsPro($worksheetIndex) as $y => $row) {
$line = [];
foreach ($row as $x => $c) {
$value = $c['value'];
if ($fillMerged && $c['merge'] && !$c['merge']['master']) {
$value = isset($mergeValues[$c['merge']['ref']]) ? $mergeValues[$c['merge']['ref']] : $value;
}
if ($fillMerged && $c['merge'] && $c['merge']['master']) {
$mergeValues[$c['merge']['ref']] = $value;
}
$cell = [
'cell'  => $c['name'],
'value' => $value,
];
if ($c['img'] !== '') {
$cell['img'] = $c['img'];
}
if ($c['f'] !== '') {
$cell['formula'] = $c['f'];
}
if ($c['calculated'] !== null) {
$cell['calculated'] = $c['calculated'];
}
if ($c['merge']) {
$cell['merge'] = [
'ref'    => $c['merge']['ref'],
'w'      => $c['merge']['w'],
'h'      => $c['merge']['h'],
'master' => $c['merge']['master'],
];
}
if (!empty($c['href'])) {
$cell['href'] = $c['href'];
}
$line[] = $cell;
}
$out[] = $line;
}
// 合并拆分：非主格回填主格值
if ($fillMerged) {
foreach ($this->mergeCells($worksheetIndex) as $m) {
$master = isset($out[$m['y1']][$m['x1']]['value']) ? $out[$m['y1']][$m['x1']]['value'] : '';
for ($y = $m['y1']; $y <= $m['y2']; $y++) {
for ($x = $m['x1']; $x <= $m['x2']; $x++) {
if (isset($out[$y][$x])) {
$out[$y][$x]['value'] = $master;
}
}
}
}
}
return $out;
}
public function saveJSON($path, $worksheetIndex = 0, array $options = [])
{
return file_put_contents($path, $this->toJSON($worksheetIndex, $options)) !== false;
}
// ------------------------------------------------------------------ 导出：HTML
/**
* HTML 导出：保留合并单元格（rowspan / colspan）+ 图片
*
* 选项：
*   images  true   是否输出 <img>
*   head    true   第一行用 <th>
*   class   表格 class
*   limit   0 行数限制
*   escape  true   htmlspecialchars
*   img_height 图片显示高度，0 为原始尺寸
*
* @param int   $worksheetIndex
* @param array $options
*
* @return string
*/
public function toHTMLMerged($worksheetIndex = 0, array $options = [])
{
$opt = array_merge([
'images'     => true,
'head'       => true,
'class'      => 'excel',
'limit'      => 0,
'escape'     => true,
'img_height' => 60,
], $options);
$this->mergeCells($worksheetIndex);
$s = '<table class="' . htmlspecialchars($opt['class'], ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '" border="1" '
. 'cellpadding="3" style="border-collapse:collapse">' . "\n";
$y = 0;
foreach ($this->readRowsPro($worksheetIndex, (int)$opt['limit']) as $row) {
$s .= "<tr>\n";
$x = 0;
foreach ($row as $c) {
$merge = isset($this->mergeMap[$worksheetIndex][$y][$x])
? $this->mergeMap[$worksheetIndex][$y][$x] : null;
// 被合并吃掉的非主格：跳过
if ($merge && !$merge['master']) {
$x++;
continue;
}
$tag = ($opt['head'] && $y === 0) ? 'th' : 'td';
$attrs = '';
if ($merge) {
if ($merge['w'] > 1) {
$attrs .= ' colspan="' . $merge['w'] . '"';
}
if ($merge['h'] > 1) {
$attrs .= ' rowspan="' . $merge['h'] . '"';
}
}
if (!empty($c['css'])) {
$attrs .= ' style="' . $c['css'] . '"';
}
$v = (string)$c['value'];
if ($opt['escape']) {
$v = htmlspecialchars($v, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}
// \R 必须加 u 修饰符：按字节匹配时会把含 0x85 等字节的
		// 多字节字符（如 UTF-8 汉字）从中间劈开，产生损坏内容
		$v = preg_replace('/\R/u', "<br>\n", $v);
$html = $v;
if ($opt['images'] && $c['img'] !== '') {
$style = $opt['img_height'] > 0 ? ' style="height:' . (int)$opt['img_height'] . 'px"' : '';
$html .= '<br><img src="' . htmlspecialchars($c['img'], ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '"' . $style
. ' alt="' . htmlspecialchars(basename($c['img']), ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8') . '">';
}
if ($html === '') {
$html = '&nbsp;';
}
$s .= '<' . $tag . $attrs . '>' . $html . '</' . $tag . ">\n";
$x++;
}
$s .= "</tr>\n";
$y++;
}
$s .= '</table>';
return $s;
}
// ------------------------------------------------------------------ 便捷：一次性转换
/**
* 一步到位：解析 + 提图 + 转换
*
* @param string $file       xlsx 路径
* @param string $format     csv | tsv | json | html | html-merged | array
* @param array  $options    ['sheet'=>0, 'img_dir'=>'Img', 'fill_merged'=>true, ...]
*
* @return array|string|false
*/
public static function convert($file, $format = 'json', array $options = [])
{
$opt = array_merge([
'sheet'       => 0,
'img_dir'     => 'Img',
'extract_img' => true,
'fill_merged' => true,
], $options);
$xlsx = self::parse($file);
if (!$xlsx) {
return false;
}
if ($opt['extract_img'] && $opt['img_dir']) {
$xlsx->imageDir = $opt['img_dir'];
$xlsx->extractImages($opt['img_dir']);
}
switch (strtolower($format)) {
case 'csv':
return $xlsx->toCSV($opt['sheet'], $opt['fill_merged']);
case 'tsv':
return $xlsx->toTSV($opt['sheet'], $opt['fill_merged']);
case 'json':
return $xlsx->toJSON($opt['sheet'], array_intersect_key($opt, array_flip([
'fill_merged', 'header', 'cells', 'meta', 'pretty', 'flags'
])));
case 'html':
return $xlsx->toHTMLEx($opt['sheet']);
case 'html-merged':
case 'htmlmerged':
return $xlsx->toHTMLMerged($opt['sheet'], array_intersect_key($opt, array_flip([
'images', 'head', 'class', 'limit', 'escape', 'img_height'
])));
}
return $xlsx->toArray($opt['sheet'], $opt['fill_merged']);
}
}