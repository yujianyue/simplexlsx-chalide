<?php

/**
 * XLSXFormula - SimpleXLSX 的轻量公式计算引擎
 *
 * 纯 PHP 实现，无外部依赖。支持：
 *   - 四则运算、幂、取余、字符串连接、比较运算
 *   - 单元格引用 / 区域引用 / 跨工作表引用（Sheet2!A1、'我的 表'!A1）
 *   - 60+ 个常用函数（数学、统计、逻辑、文本、查找、日期）
 *   - 错误值传播（#DIV/0!、#VALUE!、#N/A、#REF!、#NAME?、#NUM!）
 *   - 循环引用检测（#CIRC!）
 *   - 公式链递归计算 + 结果缓存
 *
 * 用法：
 *   $f = new XLSXFormula($provider);   // $provider($sheetIndex, 'B2') => ['value'=>..,'formula'=>..]
 *   echo $f->calculate('SUM(A1:A10)*2');
 *
 * @package Shuchkin
 * @author  SimpleXLSX Pro extension
 */

namespace Shuchkin;

class XLSXFormulaError
{
    /** @var string 形如 #DIV/0! */
    public $code;

    public function __construct($code = '#VALUE!')
    {
        $this->code = $code;
    }

    public function __toString()
    {
        return $this->code;
    }
}

/**
 * 区域引用的求值结果，函数可通过 flatten() 展开成一维数组
 */
class XLSXRange
{
    /** @var array 二维数组 [rowIndex][colIndex] => value */
    public $rows;

    public function __construct(array $rows)
    {
        $this->rows = $rows;
    }

    public function flatten()
    {
        $out = [];
        foreach ($this->rows as $row) {
            foreach ((array)$row as $v) {
                $out[] = $v;
            }
        }

        return $out;
    }

    public function count()
    {
        return count($this->flatten());
    }

    public function row($i)
    {
        return isset($this->rows[$i]) ? (array)$this->rows[$i] : [];
    }

    public function col($i)
    {
        $out = [];
        foreach ($this->rows as $row) {
            $row = (array)$row;
            $out[] = isset($row[$i]) ? $row[$i] : null;
        }

        return $out;
    }
}

class XLSXFormula
{
    /** @var callable|null function($sheetIndex, $coord): array|null */
    protected $provider;

    /** @var int 当前工作表索引 */
    protected $sheetIndex = 0;

    /** @var array 工作表名 => 索引 */
    protected $sheetNames = [];

    /** @var array 已计算过的单元格结果缓存 */
    protected $cache = [];

    /** @var array 递归调用栈，用于循环引用检测 */
    protected $stack = [];

    /** @var bool 是否把未识别的函数当成 0 而不是报错 */
    public $lenientUnknownFunctions = true;

    /** @var int 最大递归深度 */
    public $maxDepth = 64;

    /**
     * @param callable|null $provider  取值回调 function($sheetIndex, $coord)
     *                                 返回 ['value'=>mixed, 'formula'=>string] 或 null
     * @param int           $sheetIndex
     */
    public function __construct($provider = null, $sheetIndex = 0)
    {
        $this->provider = $provider;
        $this->sheetIndex = $sheetIndex;
    }

    public function setProvider(callable $provider)
    {
        $this->provider = $provider;

        return $this;
    }

    public function setSheetNames(array $names)
    {
        foreach ($names as $i => $name) {
            $this->sheetNames[self::normalizeSheetName($name)] = $i;
        }

        return $this;
    }

    public function setSheetIndex($index)
    {
        $this->sheetIndex = (int)$index;

        return $this;
    }

    public function clearCache()
    {
        $this->cache = [];
        $this->stack = [];

        return $this;
    }

    protected static function normalizeSheetName($name)
    {
        return strtolower(trim((string)$name));
    }

    /**
     * 计算单元格的值（带公式递归与缓存）
     *
     * @param string   $coord      A1
     * @param int|null $sheetIndex
     *
     * @return mixed
     */
    public function cellValue($coord, $sheetIndex = null)
    {
        $sheetIndex = $sheetIndex === null ? $this->sheetIndex : (int)$sheetIndex;
        $key = $sheetIndex . '!' . strtoupper($coord);

        if (array_key_exists($key, $this->cache)) {
            return $this->cache[$key];
        }
        if (isset($this->stack[$key])) {
            return new XLSXFormulaError('#CIRC!');
        }
        if (count($this->stack) > $this->maxDepth) {
            return new XLSXFormulaError('#REF!');
        }

        $this->stack[$key] = true;

        $data = null;
        if ($this->provider) {
            $cb = $this->provider;
            $data = $cb($sheetIndex, strtoupper($coord));
        }

        if ($data === null) {
            $value = null;
        } elseif (!empty($data['formula'])) {
            $value = $this->calculate((string)$data['formula'], $sheetIndex, $coord);
        } else {
            $value = isset($data['value']) ? $data['value'] : null;
        }

        unset($this->stack[$key]);
        $this->cache[$key] = $value;

        return $value;
    }

    /**
     * 计算公式字符串（可带等号前缀）
     *
     * @param string      $formula
     * @param int|null    $sheetIndex 公式所在工作表
     * @param string|null $coord      公式所在单元格（仅用于缓存键）
     *
     * @return mixed
     */
    public function calculate($formula, $sheetIndex = null, $coord = null)
    {
        $formula = trim((string)$formula);
        if ($formula === '') {
            return null;
        }
        if ($formula[0] === '=') {
            $formula = substr($formula, 1);
        }
        if ($sheetIndex !== null) {
            $this->sheetIndex = (int)$sheetIndex;
        }

        try {
            $tokens = $this->tokenize($formula);
            $this->pos = 0;
            $this->tokens = $tokens;
            $node = $this->parseCompare();
            if ($this->pos < count($this->tokens)) {
                // 解析未完成，说明遇到不支持的语法
                return new XLSXFormulaError('#NAME?');
            }

            return $this->evalNode($node);
        } catch (\Throwable $e) {
            return new XLSXFormulaError('#VALUE!');
        }
    }

    // ------------------------------------------------------------------ 词法分析

    protected $tokens = [];
    protected $pos = 0;

    protected function tokenize($s)
    {
        $tokens = [];
        $i = 0;
        $n = strlen($s);

        // 工作表名支持 Unicode（中文表名）与单引号包裹；非 UTF-8 时退回 ASCII 版本
        $uni = (preg_match('//u', $s) === 1);
        $reSheetRef = $uni
            ? "/^(?:'(?:[^']|'')*'|[A-Za-z_\\\\\\p{L}][\\w.\\p{L}\\p{N}\\\\ ]*)!\\s*\\\$?([A-Za-z]{1,3})\\\$?(\\d{1,7})/u"
            : "/^(?:'(?:[^']|'')*'|[A-Za-z_\\\\][\\w.\\\\ ]*)!\\s*\\\$?([A-Za-z]{1,3})\\\$?(\\d{1,7})/";
        $reRef = '/^\$?([A-Za-z]{1,3})\$?(\d{1,7})/';
        $reName = $uni
            ? "/^[A-Za-z_\\\\\\p{L}][A-Za-z0-9_.\\\\\\p{L}\\p{N}]*/u"
            : "/^[A-Za-z_\\\\][A-Za-z0-9_.\\\\]*/";
        $reQuoted = "/^'(?:[^']|'')*'/";
        $reNum = '/^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/';

        while ($i < $n) {
            $ch = $s[$i];

            if ($ch === ' ' || $ch === "\t" || $ch === "\r" || $ch === "\n") {
                $i++;
                continue;
            }

            // 字符串常量
            if ($ch === '"') {
                $j = $i + 1;
                $buf = '';
                while ($j < $n) {
                    if ($s[$j] === '"') {
                        if ($j + 1 < $n && $s[$j + 1] === '"') {
                            $buf .= '"';
                            $j += 2;
                            continue;
                        }
                        $j++;
                        break;
                    }
                    $buf .= $s[$j];
                    $j++;
                }
                $tokens[] = ['str', $buf];
                $i = $j;
                continue;
            }

            $rest = substr($s, $i);

            // 错误字面量 #DIV/0! 等
            if ($ch === '#') {
                if (preg_match('/^#(?:DIV\/0|N\/A|NUM|NAME|NULL|REF|VALUE|GETTING_DATA)[!?]?/i', $rest, $m)) {
                    $tokens[] = ['err', strtoupper($m[0])];
                    $i += strlen($m[0]);
                    continue;
                }
            }

            // 跨表引用 Sheet1!A1 / '我的 表'!A1
            if (preg_match($reSheetRef, $rest, $m)) {
                $sheet = $m[0];
                $sheet = substr($sheet, 0, strpos($sheet, '!'));
                if ($sheet !== '' && $sheet[0] === "'") {
                    $sheet = str_replace("''", "'", substr($sheet, 1, -1));
                }
                $tokens[] = ['sheetref', $sheet, self::colToIndex($m[1]), ((int)$m[2]) - 1];
                $i += strlen($m[0]);
                continue;
            }

            // 数字
            if (preg_match($reNum, $rest, $m)) {
                $tokens[] = ['num', (float)$m[0]];
                $i += strlen($m[0]);
                continue;
            }

            // 单元格引用 A1 / $A$1
            if (preg_match($reRef, $rest, $m)) {
                $tokens[] = ['ref', null, self::colToIndex($m[1]), ((int)$m[2]) - 1];
                $i += strlen($m[0]);
                continue;
            }

            // 带引号的名称（工作表名或定义名）
            if ($ch === "'" && preg_match($reQuoted, $rest, $m)) {
                $tokens[] = ['name', str_replace("''", "'", substr($m[0], 1, -1))];
                $i += strlen($m[0]);
                continue;
            }

            // 标识符：函数名 / TRUE / FALSE / 定义名
            if (preg_match($reName, $rest, $m)) {
                $tokens[] = ['name', $m[0]];
                $i += strlen($m[0]);
                continue;
            }

            // 运算符
            if (preg_match('/^(<=|>=|<>|[-+*\/^&<>=])/', $rest, $m)) {
                $tokens[] = ['op', $m[1]];
                $i += strlen($m[1]);
                continue;
            }

            // 单字符标点
            if (strpos('(),:%{};', $ch) !== false) {
                $tokens[] = ['punct', $ch];
                $i++;
                continue;
            }

            // 无法识别的字符
            $tokens[] = ['punct', '?'];
            $i++;
        }

        return $tokens;
    }

    public static function colToIndex($letters)
    {
        $letters = strtoupper(preg_replace('/[^A-Za-z]/', '', $letters));
        $n = 0;
        $len = strlen($letters);
        for ($i = 0; $i < $len; $i++) {
            $n = $n * 26 + (ord($letters[$i]) - 64);
        }

        return $n - 1;
    }

    public static function indexToCol($index)
    {
        $s = '';
        $index = (int)$index;
        while ($index >= 0) {
            $s = chr($index % 26 + 65) . $s;
            $index = (int)($index / 26) - 1;
        }

        return $s;
    }

    // ------------------------------------------------------------------ 语法分析

    protected function peek($offset = 0)
    {
        return isset($this->tokens[$this->pos + $offset]) ? $this->tokens[$this->pos + $offset] : null;
    }

    protected function next()
    {
        return isset($this->tokens[$this->pos]) ? $this->tokens[$this->pos++] : null;
    }

    protected function isOp($op, $offset = 0)
    {
        $t = $this->peek($offset);

        return $t && $t[0] === 'op' && $t[1] === $op;
    }

    protected function isPunct($p, $offset = 0)
    {
        $t = $this->peek($offset);

        return $t && $t[0] === 'punct' && $t[1] === $p;
    }

    protected function parseCompare()
    {
        $left = $this->parseConcat();
        while ($this->isOp('=') || $this->isOp('<>') || $this->isOp('<') || $this->isOp('>')
            || $this->isOp('<=') || $this->isOp('>=')) {
            $op = $this->next()[1];
            $right = $this->parseConcat();
            $left = ['bin', $op, $left, $right];
        }

        return $left;
    }

    protected function parseConcat()
    {
        $left = $this->parseAddSub();
        while ($this->isOp('&')) {
            $this->next();
            $right = $this->parseAddSub();
            $left = ['bin', '&', $left, $right];
        }

        return $left;
    }

    protected function parseAddSub()
    {
        $left = $this->parseMulDiv();
        while ($this->isOp('+') || $this->isOp('-')) {
            $op = $this->next()[1];
            $right = $this->parseMulDiv();
            $left = ['bin', $op, $left, $right];
        }

        return $left;
    }

    protected function parseMulDiv()
    {
        $left = $this->parsePower();
        while ($this->isOp('*') || $this->isOp('/')) {
            $op = $this->next()[1];
            $right = $this->parsePower();
            $left = ['bin', $op, $left, $right];
        }

        return $left;
    }

    protected function parsePower()
    {
        $left = $this->parseUnary();
        while ($this->isOp('^')) {
            $this->next();
            $right = $this->parseUnary();
            $left = ['bin', '^', $left, $right];
        }

        return $left;
    }

    protected function parseUnary()
    {
        if ($this->isOp('-') || $this->isOp('+')) {
            $op = $this->next()[1];

            return ['un', $op, $this->parseUnary()];
        }

        return $this->parsePostfix();
    }

    protected function parsePostfix()
    {
        $node = $this->parsePrimary();
        while ($this->isPunct('%')) {
            $this->next();
            $node = ['pct', $node];
        }

        return $node;
    }

    protected function parsePrimary()
    {
        $t = $this->peek();
        if ($t === null) {
            throw new \RuntimeException('unexpected end of formula');
        }

        switch ($t[0]) {
            case 'num':
                $this->next();

                return ['num', $t[1]];
            case 'str':
                $this->next();

                return ['str', $t[1]];
            case 'err':
                $this->next();

                return ['err', $t[1]];
            case 'punct':
                if ($t[1] === '(') {
                    $this->next();
                    $node = $this->parseCompare();
                    if (!$this->isPunct(')')) {
                        throw new \RuntimeException('missing )');
                    }
                    $this->next();

                    return $node;
                }
                if ($t[1] === '{') { // 数组常量 {1,2;3,4}
                    return $this->parseArrayLiteral();
                }
                break;
            case 'ref':
            case 'sheetref':
                return $this->parseReference();
            case 'name':
                $this->next();
                $name = $t[1];
                $upper = strtoupper($name);
                if ($upper === 'TRUE') {
                    return ['bool', true];
                }
                if ($upper === 'FALSE') {
                    return ['bool', false];
                }
                if ($this->isPunct('(')) {
                    $this->next();
                    $args = [];
                    if (!$this->isPunct(')')) {
                        do {
                            $args[] = $this->parseCompare();
                        } while ($this->isPunct(',') && $this->next());
                    }
                    if (!$this->isPunct(')')) {
                        throw new \RuntimeException('missing ) after function args');
                    }
                    $this->next();

                    return ['call', $upper, $args];
                }
                if ($this->isOp('!')) { // 不该出现，防御
                    break;
                }

                return ['name', $name];
        }

        throw new \RuntimeException('unexpected token ' . $t[0]);
    }

    protected function parseArrayLiteral()
    {
        $this->next(); // {
        $rows = [];
        $row = [];
        while (true) {
            if ($this->isPunct('}')) {
                $this->next();
                $rows[] = $row;
                break;
            }
            if ($this->isPunct(';')) {
                $this->next();
                $rows[] = $row;
                $row = [];
                continue;
            }
            if ($this->isPunct(',')) {
                $this->next();
                continue;
            }
            $row[] = $this->parseCompare();
        }

        return ['arr', $rows];
    }

    protected function parseReference()
    {
        $t = $this->next(); // ref / sheetref
        $sheet = $t[1];
        $c1 = $t[2];
        $r1 = $t[3];

        if ($this->isPunct(':')) {
            $this->next();
            $t2 = $this->peek();
            if ($t2 === null || ($t2[0] !== 'ref' && $t2[0] !== 'sheetref')) {
                throw new \RuntimeException('bad range');
            }
            $this->next();

            return ['range', $sheet, $c1, $r1, $t2[2], $t2[3]];
        }

        return ['ref', $sheet, $c1, $r1];
    }

    // ------------------------------------------------------------------ 求值

    protected function evalNode($node)
    {
        switch ($node[0]) {
            case 'num':
            case 'str':
            case 'bool':
                return $node[1];
            case 'err':
                return new XLSXFormulaError($node[1]);
            case 'ref':
                return $this->cellValue(self::indexToCol($node[2]) . ($node[3] + 1), $this->resolveSheet($node[1]));
            case 'range':
                return $this->evalRange($node);
            case 'arr':
                $rows = [];
                foreach ($node[1] as $row) {
                    $r = [];
                    foreach ($row as $item) {
                        $r[] = $this->evalNode($item);
                    }
                    $rows[] = $r;
                }

                return new XLSXRange($rows);
            case 'un':
                return $this->evalUnary($node[1], $this->evalNode($node[2]));
            case 'pct':
                $v = $this->toNumber($this->evalNode($node[1]));
                if ($v instanceof XLSXFormulaError) {
                    return $v;
                }

                return $v / 100;
            case 'bin':
                return $this->evalBinary($node[1], $node[2], $node[3]);
            case 'call':
                return $this->evalFunction($node[1], $node[2]);
            case 'name':
                return new XLSXFormulaError('#NAME?');
        }

        return null;
    }

    protected function resolveSheet($sheet)
    {
        if ($sheet === null) {
            return $this->sheetIndex;
        }
        if (is_numeric($sheet)) {
            return (int)$sheet;
        }
        $key = self::normalizeSheetName($sheet);

        return isset($this->sheetNames[$key]) ? $this->sheetNames[$key] : $this->sheetIndex;
    }

    protected function evalRange($node)
    {
        $sheet = $this->resolveSheet($node[1]);
        $c1 = min($node[2], $node[4]);
        $c2 = max($node[2], $node[4]);
        $r1 = min($node[3], $node[5]);
        $r2 = max($node[3], $node[5]);

        // 整列引用 A:A → 限制在一个合理范围
        $isFullCol = ($r1 === 0 && $r2 >= 1048575);
        if ($isFullCol) {
            $r2 = $r1 + 1000;
        }

        $rows = [];
        for ($r = $r1; $r <= $r2; $r++) {
            $row = [];
            for ($c = $c1; $c <= $c2; $c++) {
                $row[] = $this->cellValue(self::indexToCol($c) . ($r + 1), $sheet);
            }
            $rows[] = $row;
        }

        return new XLSXRange($rows);
    }

    protected function evalUnary($op, $v)
    {
        if ($v instanceof XLSXFormulaError) {
            return $v;
        }
        if ($op === '+') {
            return $this->toNumber($v);
        }
        $n = $this->toNumber($v);
        if ($n instanceof XLSXFormulaError) {
            return $n;
        }

        return -$n;
    }

    /**
     * 二元运算。$leftNode/$rightNode 是 AST，按需求值（用于短路）
     */
    protected function evalBinary($op, $leftNode, $rightNode)
    {
        $l = $this->evalNode($leftNode);
        if ($l instanceof XLSXFormulaError) {
            return $l;
        }
        $r = $this->evalNode($rightNode);
        if ($r instanceof XLSXFormulaError) {
            return $r;
        }

        if ($op === '&') {
            return $this->toString($l) . $this->toString($r);
        }

        if (in_array($op, ['=', '<>', '<', '>', '<=', '>='], true)) {
            return $this->compare($l, $r, $op);
        }

        $a = $this->toNumber($l);
        if ($a instanceof XLSXFormulaError) {
            return $a;
        }
        $b = $this->toNumber($r);
        if ($b instanceof XLSXFormulaError) {
            return $b;
        }

        switch ($op) {
            case '+':
                return $a + $b;
            case '-':
                return $a - $b;
            case '*':
                return $a * $b;
            case '/':
                if ($b == 0) {
                    return new XLSXFormulaError('#DIV/0!');
                }

                return $a / $b;
            case '^':
                $res = pow($a, $b);
                if (is_nan($res) || is_infinite($res)) {
                    return new XLSXFormulaError('#NUM!');
                }

                return $res;
        }

        return new XLSXFormulaError('#VALUE!');
    }

    protected function compare($l, $r, $op)
    {
        // 两边都是数字，或其中一个是数字型字符串 → 数值比较
        $ln = $this->isNumeric($l);
        $rn = $this->isNumeric($r);

        if ($ln && $rn) {
            $a = $this->toNumber($l);
            $b = $this->toNumber($r);
        } else {
            $a = $this->toString($l);
            $b = $this->toString($r);
            if (!$ln && !$rn && is_string($a) && is_string($b)) {
                $a = strtolower($a);
                $b = strtolower($b);
            }
        }

        switch ($op) {
            case '=':
                $res = ($a == $b);
                break;
            case '<>':
                $res = ($a != $b);
                break;
            case '<':
                $res = ($a < $b);
                break;
            case '>':
                $res = ($a > $b);
                break;
            case '<=':
                $res = ($a <= $b);
                break;
            case '>=':
                $res = ($a >= $b);
                break;
            default:
                $res = false;
        }

        return $res;
    }

    // ------------------------------------------------------------------ 类型转换

    public static function isNumeric($v)
    {
        if ($v === null || $v === '' || $v === false) {
            return false;
        }
        if (is_bool($v)) {
            return false;
        }
        if (is_int($v) || is_float($v)) {
            return true;
        }
        if ($v instanceof XLSXFormulaError || $v instanceof XLSXRange) {
            return false;
        }
        if (is_string($v)) {
            $v = trim($v);
            if ($v === '') {
                return false;
            }
            // 去掉千分位与百分号
            $v = str_replace([' ', ','], '', $v);
            if (substr($v, -1) === '%') {
                $v = substr($v, 0, -1);
            }

            return is_numeric($v);
        }

        return false;
    }

    /**
     * 转数字，失败返回 XLSXFormulaError
     */
    public function toNumber($v)
    {
        if ($v instanceof XLSXFormulaError) {
            return $v;
        }
        if ($v === null || $v === '') {
            return 0;
        }
        if (is_bool($v)) {
            return $v ? 1 : 0;
        }
        if (is_int($v) || is_float($v)) {
            return $v;
        }
        if ($v instanceof XLSXRange) {
            $f = $v->flatten();

            return isset($f[0]) ? $this->toNumber($f[0]) : 0;
        }
        if (is_string($v)) {
            $s = trim($v);
            $s = str_replace([' ', ','], '', $s);
            $pct = false;
            if (substr($s, -1) === '%') {
                $pct = true;
                $s = substr($s, 0, -1);
            }
            if (is_numeric($s)) {
                $n = $s + 0;

                return $pct ? $n / 100 : $n;
            }
            // 尝试日期字符串
            $ts = strtotime($v);
            if ($ts !== false) {
                return self::dateToSerial($ts);
            }

            return new XLSXFormulaError('#VALUE!');
        }

        return new XLSXFormulaError('#VALUE!');
    }

    public function toString($v)
    {
        if ($v instanceof XLSXFormulaError) {
            return $v->code;
        }
        if ($v === null) {
            return '';
        }
        if (is_bool($v)) {
            return $v ? 'TRUE' : 'FALSE';
        }
        if (is_float($v)) {
            if ($v == (int)$v && abs($v) < 1e15) {
                return (string)(int)$v;
            }

            return (string)$v;
        }
        if (is_array($v)) {
            return '';
        }
        if ($v instanceof XLSXRange) {
            $f = $v->flatten();

            return isset($f[0]) ? $this->toString($f[0]) : '';
        }

        return (string)$v;
    }

    /** Excel 日期序列号：1900 系统，含 1900 闰年 bug */
    public static function dateToSerial($ts)
    {
        return ($ts / 86400) + 25569;
    }

    public static function serialToTs($serial)
    {
        return (int)round(($serial - 25569) * 86400);
    }

    // ------------------------------------------------------------------ 函数库

    protected function evalFunction($name, array $argNodes)
    {
        $name = strtoupper($name);

        // 需要惰性求值的特殊函数
        switch ($name) {
            case 'IF':
                $cond = $this->evalNode($argNodes[0]);
                if ($cond instanceof XLSXFormulaError) {
                    return $cond;
                }
                $bool = $this->toBool($cond);
                if ($bool) {
                    return isset($argNodes[1]) ? $this->evalNode($argNodes[1]) : true;
                }

                return isset($argNodes[2]) ? $this->evalNode($argNodes[2]) : false;

            case 'IFERROR':
                $v = $this->evalNode($argNodes[0]);
                if ($v instanceof XLSXFormulaError) {
                    return isset($argNodes[1]) ? $this->evalNode($argNodes[1]) : '';
                }

                return $v;

            case 'IFNA':
                $v = $this->evalNode($argNodes[0]);
                if ($v instanceof XLSXFormulaError && $v->code === '#N/A') {
                    return isset($argNodes[1]) ? $this->evalNode($argNodes[1]) : '';
                }

                return $v;

            case 'AND':
            case 'OR':
                $result = ($name === 'AND');
                foreach ($argNodes as $an) {
                    $v = $this->evalNode($an);
                    if ($v instanceof XLSXFormulaError) {
                        return $v;
                    }
                    $b = $this->toBool($v);
                    if ($name === 'AND' && !$b) {
                        $result = false;
                    } elseif ($name === 'OR' && $b) {
                        $result = true;
                    }
                }

                return $result;

            case 'VLOOKUP':
                return $this->fnVLookup($argNodes, false);
            case 'HLOOKUP':
                return $this->fnVLookup($argNodes, true);
            case 'INDEX':
                return $this->fnIndex($argNodes);
            case 'MATCH':
                return $this->fnMatch($argNodes);
            case 'OFFSET':
                return $this->fnOffset($argNodes);
        }

        // 其余函数：先求值参数
        $args = [];
        foreach ($argNodes as $an) {
            $args[] = $this->evalNode($an);
        }
        foreach ($args as $a) {
            if ($a instanceof XLSXFormulaError) {
                return $a; // 错误传播
            }
        }

        $a = isset($args[0]) ? $args[0] : null;
        $b = isset($args[1]) ? $args[1] : null;
        $c = isset($args[2]) ? $args[2] : null;

        switch ($name) {
            // ---------- 数学 / 统计 ----------
            case 'SUM':
                return $this->sumNumbers($args);
            case 'SUMIF':
                return $this->fnSumIf($args);
            case 'SUMPRODUCT':
                return $this->fnSumProduct($args);
            case 'PRODUCT':
                $p = 1;
                foreach ($this->numbers($args) as $n) {
                    $p *= $n;
                }

                return $p;
            case 'AVERAGE':
                $n = $this->numbers($args);
                if (!count($n)) {
                    return new XLSXFormulaError('#DIV/0!');
                }

                return array_sum($n) / count($n);
            case 'MEDIAN':
                $n = $this->numbers($args);
                if (!count($n)) {
                    return new XLSXFormulaError('#NUM!');
                }
                sort($n);
                $m = count($n);

                return $m % 2 ? $n[(int)($m / 2)] : ($n[$m / 2 - 1] + $n[$m / 2]) / 2;
            case 'MAX':
            case 'MIN':
                $n = $this->numbers($args);
                if (!count($n)) {
                    return 0;
                }

                return $name === 'MAX' ? max($n) : min($n);
            case 'COUNT':
                $k = 0;
                foreach ($args as $arg) {
                    $vals = $arg instanceof XLSXRange ? $arg->flatten() : [$arg];
                    foreach ($vals as $v) {
                        if (is_int($v) || is_float($v)) {
                            $k++;
                        }
                    }
                }

                return $k;
            case 'COUNTA':
                $k = 0;
                foreach ($args as $arg) {
                    $vals = $arg instanceof XLSXRange ? $arg->flatten() : [$arg];
                    foreach ($vals as $v) {
                        if ($v !== null && $v !== '') {
                            $k++;
                        }
                    }
                }

                return $k;
            case 'COUNTBLANK':
                $k = 0;
                foreach ($args as $arg) {
                    $vals = $arg instanceof XLSXRange ? $arg->flatten() : [$arg];
                    foreach ($vals as $v) {
                        if ($v === null || $v === '') {
                            $k++;
                        }
                    }
                }

                return $k;
            case 'COUNTIF':
                return $this->fnCountIf($args);
            case 'ROUND':
                return round($this->num($a), (int)$this->num($b === null ? 0 : $b));
            case 'ROUNDUP':
                $p = pow(10, (int)$this->num($b === null ? 0 : $b));
                $v = $this->num($a) * $p;

                return ($v >= 0 ? ceil($v) : floor($v)) / $p;
            case 'ROUNDDOWN':
                $p = pow(10, (int)$this->num($b === null ? 0 : $b));
                $v = $this->num($a) * $p;

                return ($v >= 0 ? floor($v) : ceil($v)) / $p;
            case 'TRUNC':
                $p = pow(10, (int)$this->num($b === null ? 0 : $b));

                return (int)($this->num($a) * $p) / $p;
            case 'INT':
                return (int)floor($this->num($a));
            case 'ABS':
                return abs($this->num($a));
            case 'SIGN':
                $n = $this->num($a);

                return $n > 0 ? 1 : ($n < 0 ? -1 : 0);
            case 'CEILING':
                $s = $this->num($b === null ? 1 : $b);

                return $s == 0 ? 0 : ceil($this->num($a) / $s) * $s;
            case 'FLOOR':
                $s = $this->num($b === null ? 1 : $b);

                return $s == 0 ? 0 : floor($this->num($a) / $s) * $s;
            case 'MOD':
                $d = $this->num($b);
                if ($d == 0) {
                    return new XLSXFormulaError('#DIV/0!');
                }

                return $this->num($a) - $d * (int)floor($this->num($a) / $d);
            case 'POWER':
                return pow($this->num($a), $this->num($b));
            case 'SQRT':
                $n = $this->num($a);

                return $n < 0 ? new XLSXFormulaError('#NUM!') : sqrt($n);
            case 'EXP':
                return exp($this->num($a));
            case 'LN':
                $n = $this->num($a);

                return $n <= 0 ? new XLSXFormulaError('#NUM!') : log($n);
            case 'LOG':
                $base = $b === null ? 10 : $this->num($b);

                return log($this->num($a), $base);
            case 'LOG10':
                return log10($this->num($a));
            case 'PI':
                return M_PI;
            case 'RAND':
                return mt_rand() / mt_getrandmax();
            case 'RANDBETWEEN':
                return mt_rand((int)$this->num($a), (int)$this->num($b));
            case 'DEGREES':
                return rad2deg($this->num($a));
            case 'RADIANS':
                return deg2rad($this->num($a));
            case 'STDEV':
            case 'STDEVP':
                $n = $this->numbers($args);
                $c = count($n);
                if ($c < 2) {
                    return 0;
                }
                $mean = array_sum($n) / $c;
                $s = 0;
                foreach ($n as $v) {
                    $s += ($v - $mean) ** 2;
                }

                return sqrt($s / ($name === 'STDEV' ? $c - 1 : $c));
            case 'VAR':
                $n = $this->numbers($args);
                $c = count($n);
                if ($c < 2) {
                    return 0;
                }
                $mean = array_sum($n) / $c;
                $s = 0;
                foreach ($n as $v) {
                    $s += ($v - $mean) ** 2;
                }

                return $s / ($c - 1);
            case 'LARGE':
            case 'SMALL':
                $n = $this->numbers($args);
                $k = (int)$this->num($b);
                if ($name === 'LARGE') {
                    rsort($n);
                } else {
                    sort($n);
                }

                return isset($n[$k - 1]) ? $n[$k - 1] : new XLSXFormulaError('#NUM!');
            case 'SUBTOTAL':
                return $this->sumNumbers(array_slice($args, 1));

            // ---------- 逻辑 / 信息 ----------
            case 'NOT':
                return !$this->toBool($a);
            case 'TRUE':
                return true;
            case 'FALSE':
                return false;
            case 'ISBLANK':
                return ($a === null || $a === '');
            case 'ISNUMBER':
                return self::isNumeric($a);
            case 'ISTEXT':
                return is_string($a) && !self::isNumeric($a);
            case 'ISERROR':
                return ($a instanceof XLSXFormulaError);
            case 'ISNA':
                return ($a instanceof XLSXFormulaError && $a->code === '#N/A');
            case 'ISERR':
                return ($a instanceof XLSXFormulaError && $a->code !== '#N/A');
            case 'NA':
                return new XLSXFormulaError('#N/A');
            case 'N':
                return $this->toNumber($a);
            case 'T':
                return is_string($a) ? $a : '';

            // ---------- 文本 ----------
            case 'CONCATENATE':
            case 'CONCAT':
            case 'TEXTJOIN':
                if ($name === 'TEXTJOIN') {
                    $delim = $this->toString($a);
                    $skipEmpty = (bool)$this->toBool($b);
                    $parts = [];
                    foreach (array_slice($args, 2) as $arg) {
                        foreach ($this->flattenArg($arg) as $v) {
                            if ($skipEmpty && ($v === '' || $v === null)) {
                                continue;
                            }
                            $parts[] = $this->toString($v);
                        }
                    }

                    return implode($delim, $parts);
                }
                $out = '';
                foreach ($args as $arg) {
                    foreach ($this->flattenArg($arg) as $v) {
                        $out .= $this->toString($v);
                    }
                }

                return $out;
            case 'LEN':
                return strlen($this->toString($a));
            case 'LEFT':
                return substr($this->toString($a), 0, max(0, (int)$this->num($b === null ? 1 : $b)));
            case 'RIGHT':
                $s = $this->toString($a);
                $n = max(0, (int)$this->num($b === null ? 1 : $b));

                return $n === 0 ? '' : substr($s, -$n);
            case 'MID':
                return substr($this->toString($a), max(0, (int)$this->num($b) - 1), (int)$this->num($c === null ? 1 : $c));
            case 'UPPER':
                return strtoupper($this->toString($a));
            case 'LOWER':
                return strtolower($this->toString($a));
            case 'PROPER':
                return ucwords(strtolower($this->toString($a)));
            case 'TRIM':
                return trim(preg_replace('/\s+/', ' ', $this->toString($a)));
            case 'SUBSTITUTE':
                return str_replace(
                    $this->toString($b),
                    $this->toString($c),
                    $this->toString($a)
                );
            case 'REPLACE':
                // REPLACE(old_text, start_num, num_chars, new_text)
                $s = $this->toString($a);
                $start = max(1, (int)$this->num($b));
                $len = (int)$this->num($c === null ? 0 : $c);
                $new = $this->toString(isset($args[3]) ? $args[3] : '');

                return substr($s, 0, $start - 1) . $new . substr($s, $start - 1 + $len);
            case 'FIND':
            case 'SEARCH':
                $hay = $this->toString($a);
                $needle = $this->toString($b);
                $offset = max(0, (int)$this->num($c === null ? 1 : $c) - 1);
                if ($name === 'FIND') {
                    $p = strpos($hay, $needle, $offset);
                } else {
                    $p = stripos($hay, $needle, $offset);
                }
                if ($p === false) {
                    return new XLSXFormulaError('#VALUE!');
                }

                return $p + 1;
            case 'REPT':
                return str_repeat($this->toString($a), max(0, (int)$this->num($b)));
            case 'TEXT':
                return $this->fnText($a, $this->toString($b));
            case 'VALUE':
                return $this->toNumber($a);
            case 'CHAR':
                return chr((int)$this->num($a));
            case 'CODE':
                $s = $this->toString($a);

                return $s === '' ? new XLSXFormulaError('#VALUE!') : ord($s[0]);
            case 'EXACT':
                return $this->toString($a) === $this->toString($b);
            case 'SPLIT':
                return new XLSXRange([explode($this->toString($b), $this->toString($a))]);

            // ---------- 查找 ----------
            case 'CHOOSE':
                $i = (int)$this->num($a);

                return isset($args[$i]) ? $args[$i] : new XLSXFormulaError('#VALUE!');

            // ---------- 日期 ----------
            case 'TODAY':
                return self::dateToSerial(strtotime(date('Y-m-d')));
            case 'NOW':
                return self::dateToSerial(time());
            case 'DATE':
                // DATE(year, month, day)
                return self::dateToSerial(mktime(
                    0,
                    0,
                    0,
                    (int)$this->num($b),
                    (int)$this->num($c === null ? 1 : $c),
                    (int)$this->num($a)
                ));
            case 'YEAR':
                return (int)date('Y', $this->tsOf($a));
            case 'MONTH':
                return (int)date('n', $this->tsOf($a));
            case 'DAY':
                return (int)date('j', $this->tsOf($a));
            case 'HOUR':
                return (int)date('G', $this->tsOf($a));
            case 'MINUTE':
                return (int)date('i', $this->tsOf($a));
            case 'SECOND':
                return (int)date('s', $this->tsOf($a));
            case 'WEEKDAY':
                return (int)date('w', $this->tsOf($a)) + 1;
            case 'DATEVALUE':
                return $this->toNumber($a);
            case 'DAYS':
                return (int)round($this->tsOf($b) / 86400) - (int)round($this->tsOf($a) / 86400);
            case 'DATEDIF':
                $t1 = $this->tsOf($a);
                $t2 = $this->tsOf($b);
                $unit = strtolower($this->toString($c === null ? 'd' : $c));
                if ($t2 < $t1) {
                    return new XLSXFormulaError('#NUM!');
                }
                $d1 = new \DateTimeImmutable('@' . $t1);
                $d2 = new \DateTimeImmutable('@' . $t2);
                $diff = $d1->diff($d2);
                switch ($unit) {
                    case 'y':
                        return $diff->y;
                    case 'm':
                        return $diff->y * 12 + $diff->m;
                    case 'd':
                        return $diff->days;
                    case 'md':
                        return $diff->d;
                    case 'ym':
                        return $diff->m;
                    case 'yd':
                        return (int)round(($t2 - $t1) / 86400) % 365;
                }

                return $diff->days;
            case 'EDATE':
                $t = $this->tsOf($a);

                return self::dateToSerial(strtotime('+' . (int)$this->num($b) . ' months', $t));
            case 'EOMONTH':
                $t = $this->tsOf($a);
                $t = strtotime('+' . (int)$this->num($b) . ' months', $t);

                return self::dateToSerial(strtotime('last day of this month', $t));
        }

        if ($this->lenientUnknownFunctions) {
            return 0;
        }

        return new XLSXFormulaError('#NAME?');
    }

    protected function anyTrue(array $argNodes)
    {
        foreach ($argNodes as $an) {
            if ($this->toBool($this->evalNode($an))) {
                return true;
            }
        }

        return false;
    }

    protected function toBool($v)
    {
        if ($v instanceof XLSXFormulaError) {
            return false;
        }
        if (is_bool($v)) {
            return $v;
        }
        if ($v === null || $v === '') {
            return false;
        }
        if (is_string($v)) {
            $u = strtoupper(trim($v));
            if ($u === 'TRUE') {
                return true;
            }
            if ($u === 'FALSE') {
                return false;
            }

            return true;
        }

        return (bool)$v;
    }

    protected function num($v)
    {
        if ($v instanceof XLSXRange) {
            $f = $v->flatten();
            $v = isset($f[0]) ? $f[0] : 0;
        }
        if (is_bool($v)) {
            return $v ? 1 : 0;
        }
        if ($v === null || $v === '') {
            return 0;
        }
        if (is_numeric($v)) {
            return $v + 0;
        }
        if (is_string($v)) {
            $s = str_replace([' ', ','], '', trim($v));
            $pct = false;
            if (substr($s, -1) === '%') {
                $pct = true;
                $s = substr($s, 0, -1);
            }
            if (is_numeric($s)) {
                return $pct ? $s / 100 : $s + 0;
            }
            $ts = strtotime($v);
            if ($ts !== false) {
                return self::dateToSerial($ts);
            }
        }

        return 0;
    }

    protected function numbers(array $args)
    {
        $out = [];
        foreach ($args as $arg) {
            foreach ($this->flattenArg($arg) as $v) {
                if ($v instanceof XLSXFormulaError) {
                    continue;
                }
                if (is_bool($v)) {
                    continue;
                }
                if (is_int($v) || is_float($v)) {
                    $out[] = $v;
                } elseif (is_string($v) && self::isNumeric($v)) {
                    $out[] = $this->num($v);
                }
            }
        }

        return $out;
    }

    protected function sumNumbers(array $args)
    {
        return array_sum($this->numbers($args));
    }

    protected function flattenArg($arg)
    {
        if ($arg instanceof XLSXRange) {
            return $arg->flatten();
        }
        if (is_array($arg)) {
            return $arg;
        }

        return [$arg];
    }

    protected function tsOf($v)
    {
        $n = $this->num($v);
        // Excel 日期序列号的范围是 1 .. 2958465 (9999-12-31)，
        // 超出这个范围才当成 Unix 时间戳
        if ($n > 0 && $n < 2958466) {
            return self::serialToTs($n);
        }

        return (int)$n;
    }

    // ---------------------------------------------------------------- 具体函数

    protected function fnCountIf(array $args)
    {
        $range = $args[0];
        $crit = isset($args[1]) ? $args[1] : null;
        $vals = $this->flattenArg($range);
        $n = 0;
        foreach ($vals as $v) {
            if ($this->matchCriteria($v, $crit)) {
                $n++;
            }
        }

        return $n;
    }

    protected function fnSumIf(array $args)
    {
        $range = $args[0];
        $crit = isset($args[1]) ? $args[1] : null;
        $sumRange = isset($args[2]) ? $args[2] : $range;

        $vals = $this->flattenArg($range);
        $sums = $this->flattenArg($sumRange);

        $total = 0;
        foreach ($vals as $i => $v) {
            if ($this->matchCriteria($v, $crit)) {
                $total += $this->num(isset($sums[$i]) ? $sums[$i] : 0);
            }
        }

        return $total;
    }

    protected function fnSumProduct(array $args)
    {
        $lists = [];
        foreach ($args as $arg) {
            $lists[] = $this->flattenArg($arg);
        }
        if (!count($lists)) {
            return 0;
        }
        $len = count($lists[0]);
        $total = 0;
        for ($i = 0; $i < $len; $i++) {
            $p = 1;
            foreach ($lists as $list) {
                $p *= $this->num(isset($list[$i]) ? $list[$i] : 0);
            }
            $total += $p;
        }

        return $total;
    }

    /**
     * 条件匹配，支持 ">10"、"<=5"、"<>abc"、"abc"、"A*" 等
     */
    protected function matchCriteria($value, $crit)
    {
        if ($crit === null) {
            return $value === null || $value === '';
        }
        $crit = $this->toString($crit);

        if (preg_match('/^(>=|<=|<>|>|<|=)\s*(.*)$/', $crit, $m)) {
            $op = $m[1] === '=' ? '=' : $m[1];
            $target = $m[2];

            return $this->compare($value, self::isNumeric($target) ? $this->num($target) : $target, $op);
        }
        if ($value === null) {
            return false;
        }
        if (self::isNumeric($crit) && self::isNumeric($value)) {
            return $this->num($value) == $this->num($crit);
        }
        $crit = strtolower($crit);
        $s = strtolower($this->toString($value));
        if (strpos($crit, '*') !== false || strpos($crit, '?') !== false) {
            $re = '/^' . str_replace(['\*', '\?'], ['.*', '.'], preg_quote($crit, '/')) . '$/';
            $re = str_replace(['\*', '\?'], ['.*', '.'], $re);

            return (bool)preg_match($re, $s);
        }

        return $s === $crit;
    }

    protected function fnVLookup(array $argNodes, $horizontal)
    {
        $needle = $this->evalNode($argNodes[0]);
        $table = $this->evalNode($argNodes[1]);
        $colIdx = isset($argNodes[2]) ? (int)$this->num($this->evalNode($argNodes[2])) : 1;
        $exact = isset($argNodes[3]) ? !$this->toBool($this->evalNode($argNodes[3])) : true;

        if (!($table instanceof XLSXRange)) {
            $table = new XLSXRange([$this->flattenArg($table)]);
        }
        $rows = $table->rows;
        if ($horizontal) {
            $rows = $this->transpose($rows);
        }

        $best = null;
        $bestVal = null;
        foreach ($rows as $row) {
            $row = (array)$row;
            $first = isset($row[0]) ? $row[0] : null;
            if ($this->compare($first, $needle, '=')) {
                return isset($row[$colIdx - 1]) ? $row[$colIdx - 1] : new XLSXFormulaError('#REF!');
            }
            if (!$exact && self::isNumeric($first) && self::isNumeric($needle)
                && $this->num($first) <= $this->num($needle)) {
                if ($bestVal === null || $this->num($first) > $bestVal) {
                    $best = $row;
                    $bestVal = $this->num($first);
                }
            }
        }
        if ($best !== null) {
            return isset($best[$colIdx - 1]) ? $best[$colIdx - 1] : new XLSXFormulaError('#REF!');
        }

        return new XLSXFormulaError('#N/A');
    }

    protected function fnIndex(array $argNodes)
    {
        $range = $this->evalNode($argNodes[0]);
        $rowNum = isset($argNodes[1]) ? (int)$this->num($this->evalNode($argNodes[1])) : 1;
        $colNum = isset($argNodes[2]) ? (int)$this->num($this->evalNode($argNodes[2])) : null;

        if (!($range instanceof XLSXRange)) {
            $range = new XLSXRange([$this->flattenArg($range)]);
        }
        $rows = $range->rows;
        if ($colNum === null) {
            $flat = $range->flatten();

            return isset($flat[$rowNum - 1]) ? $flat[$rowNum - 1] : new XLSXFormulaError('#REF!');
        }
        $row = isset($rows[$rowNum - 1]) ? (array)$rows[$rowNum - 1] : [];

        return isset($row[$colNum - 1]) ? $row[$colNum - 1] : new XLSXFormulaError('#REF!');
    }

    protected function fnMatch(array $argNodes)
    {
        $needle = $this->evalNode($argNodes[0]);
        $range = $this->evalNode($argNodes[1]);
        $type = isset($argNodes[2]) ? (int)$this->num($this->evalNode($argNodes[2])) : 1;

        $vals = $this->flattenArg($range);
        foreach ($vals as $i => $v) {
            if ($type === 0 && $this->compare($v, $needle, '=')) {
                return $i + 1;
            }
            if ($type !== 0 && self::isNumeric($v) && self::isNumeric($needle)) {
                if ($type === 1 && $this->num($v) == $this->num($needle)) {
                    return $i + 1;
                }
                if ($type === -1 && $this->num($v) == $this->num($needle)) {
                    return $i + 1;
                }
            } elseif ($type === 0 && $this->toString($v) === $this->toString($needle)) {
                return $i + 1;
            }
        }

        return new XLSXFormulaError('#N/A');
    }

    protected function fnOffset(array $argNodes)
    {
        $base = $this->evalNode($argNodes[0]);
        if (!($base instanceof XLSXRange)) {
            return $base;
        }
        $rows = $base->rows;
        $rOff = isset($argNodes[1]) ? (int)$this->num($this->evalNode($argNodes[1])) : 0;
        $cOff = isset($argNodes[2]) ? (int)$this->num($this->evalNode($argNodes[2])) : 0;
        $h = isset($argNodes[3]) ? (int)$this->num($this->evalNode($argNodes[3])) : null;
        $w = isset($argNodes[4]) ? (int)$this->num($this->evalNode($argNodes[4])) : null;

        $out = [];
        for ($r = $rOff; $r < ($h === null ? count($rows) : $rOff + $h); $r++) {
            if (!isset($rows[$r])) {
                continue;
            }
            $row = (array)$rows[$r];
            $outRow = [];
            for ($c = $cOff; $c < ($w === null ? count($row) : $cOff + $w); $c++) {
                $outRow[] = isset($row[$c]) ? $row[$c] : null;
            }
            $out[] = $outRow;
        }

        return new XLSXRange($out);
    }

    protected function transpose(array $rows)
    {
        $out = [];
        foreach ($rows as $r => $row) {
            foreach ((array)$row as $c => $v) {
                $out[$c][$r] = $v;
            }
        }

        return $out;
    }

    /**
     * TEXT() 的最小实现，支持常见的数字与日期格式
     */
    protected function fnText($value, $format)
    {
        $n = $this->num($value);
        $fmt = trim($format);

        // 日期格式
        if (preg_match('/[yY]/', $fmt) && preg_match('/[mdhs]/i', $fmt)) {
            $ts = $n > 100000 ? self::serialToTs($n) : (int)$n;
            $map = [
                'yyyy' => 'Y', 'yy' => 'y', 'mmmm' => 'F', 'mmm' => 'M', 'mm' => 'm', 'm' => 'n',
                'dd' => 'd', 'd' => 'j', 'hh' => 'H', 'h' => 'G', 'ss' => 's', 's' => 's',
            ];
            $out = '';
            $i = 0;
            $len = strlen($fmt);
            while ($i < $len) {
                $matched = false;
                foreach ($map as $token => $php) {
                    if (strtolower(substr($fmt, $i, strlen($token))) === $token) {
                        $out .= date($php, $ts);
                        $i += strlen($token);
                        $matched = true;
                        break;
                    }
                }
                if (!$matched) {
                    $out .= $fmt[$i];
                    $i++;
                }
            }

            return $out;
        }

        // 百分比
        if (substr($fmt, -1) === '%') {
            $dec = substr_count($fmt, '0') ? substr_count($fmt, '0') - 1 : 0;
            if (strpos($fmt, '.') !== false) {
                preg_match('/\.([0#]+)/', $fmt, $m);
                $dec = isset($m[1]) ? strlen($m[1]) : 0;
            }

            return number_format($n * 100, $dec) . '%';
        }

        // 数字
        $dec = 0;
        if (strpos($fmt, '.') !== false) {
            preg_match('/\.([0#]+)/', $fmt, $m);
            $dec = isset($m[1]) ? strlen($m[1]) : 0;
        }
        $thousands = strpos($fmt, ',') !== false;

        return number_format($n, $dec, '.', $thousands ? ',' : '');
    }
}
