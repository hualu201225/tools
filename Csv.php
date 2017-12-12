<?php
/**
 * Name: 数据导出类(csv)
 * Created by PhpStorm.
 * User: hualu@myhexin.com
 * Date: 2017/11/14
 * Time: 14:50
 */

class Public_Excel_Csv
{

    /**
     * @var string
     */
    private $_enclosure = '"';

    /**
     * 分隔符
     * @var string
     */
    private $_delimiter = ',';

    /**
     * 结束符
     * @var string
     */
    private $_ending = PHP_EOL;

    /**
     * 文件后缀
     * @var string
     */
    private $_postFix = 'csv';

    /**
     * 编码
     * @var string
     */
    private $_encoding = 'GBK';

    /**
     * 文件句柄
     * @var null
     */
    private $_fileop = null;

    /**
     * 循环写入时每次写入的条数限制
     * @var int
     */
    private $_limitCount = 200;

    /**
     * 数据列标题（一维数组）
     */
    private $_head = null;

    /**
     * 数据（二维数组）
     */
    private $_data = null;

    /**
     * 文件默认路径
     */
    private $_path = '/tmp/';

    /**
     * 文件名称
     */
    private $_filename = '';

    /**
     * 文件标题标志
     * @var int
     */
    private $_headExists = 0;

    /**
     * 浏览器header设置标志
     * @var int
     */
    private $_headerExists = 0;

    /**
     * 时间戳标志
     * @var null
     */
    private $_timestamp = null;

    /**
     * 是否浏览器直接输出
     * @var int
     */
    private $_isOutput = 0;

    /**
     * 是否从文件尾写入数据
     * @var int
     */
    private $_isAppend = 1;

    /**
     * 错误代码
     */
    private $_errorCode = 0;

    /**
     * 列标题格式错误
     */
    const ERROR_HEADVALID = -1;

    /**
     * 数据格式错误
     */
    const ERROR_DATAVALID = -2;

    /**
     * 文件下载路径错误
     */
    const ERROR_PATH = -3;

    /**
     * 错误信息
     */
    private static $_errorMsgs = [
        self::ERROR_HEADVALID => '头部传入不正确(应是一维数组)',
        self::ERROR_DATAVALID => '数据传入格式不正确(应是二维数组)',
        self::ERROR_PATH => '文件路径不正确，请重新输入'
    ];

    public function setTitle($title)
    {
        $this->_title = $title;
        return $this;
    }

    public function setHead(array $head)
    {
        $this->_head = $head;
        return $this;
    }

    public function setData(array $data)
    {
        $this->_data = $data;
        return $this;
    }

    public function setFileName($filename)
    {
        $this->_filename = $filename;
        return $this;
    }

    public function setPath($path)
    {
        $this->_path = $path;
        return $this;
    }

    /**
     * 设置是否浏览器输出
     * @param bool $bool 为true时为浏览器输出
     * @create by hualu
     */
    public function setIsOut($bool = false)
    {
        $this->_isOutput = (int)$bool;
        return $this;
    }

    /**
     * 设置文件句柄
     * @param bool $isOutPut 是否直接输出到浏览器
     * @create by hualu
     */
    private function _setFileop()
    {
        //句柄不存在/追加数据时均重新获取文件句柄
        if (!$this->_fileop) {
            if (!$this->_isOutput) {
                $filepath = $this->_path . $this->_filename;
                $mode = $this->_isAppend ? 'ab+' : 'wb+';
                $this->_fileop = fopen($filepath, $mode);
            } else {
                $this->_fileop = fopen("php://output", 'wb+');
            }
        }
    }

    /**
     * 关闭文件句柄
     * @create by hualu
     */
    private function _closeFileop()
    {
        if ($this->_fileop) {
            fclose($this->_fileop);
            $this->_fileop = null;
        }
    }

    /**
     * 重启文件句柄
     * @create by hualu
     */
    private function _resetFileop()
    {
        $this->_closeFileop();
        $this->_setFileop();
    }

    /**
     * 编码转换
     * @param $content
     * @return string
     * @create by hualu
     */
    private function _getEncodingContent($content)
    {
        return mb_convert_encoding($content, $this->_encoding);
    }

    /**
     * 设置data
     * @create by hualu
     */
    private function _mergeHeadAndData()
    {
        if (!$this->_headExists) {
            array_unshift($this->_data, $this->_head);
            $this->_headExists = 1;
        }
    }

    /**
     * 写入数据（支持一、二维数组）
     * @param array $data
     * @return array
     * @create by hualu
     */
    public function write(array $data)
    {
        //初始化文件路径、名称等
        $this->_initFilePath();

        //当$data是一维数组时，将其转为二维数组
        if ($this->_isOneDimension($data)) {
            $data = [$data];
        }

        //校验传入数据
        $confirm = $this->_paramsCheck($data);
        if (!$confirm) {
            return [$this->_errorCode, self::$_errorMsgs[$this->_errorCode]];
        }

        //设置将列标题和数据整合
        $this->_mergeHeadAndData();

        //写入数据
        $this->_appendData();
    }

    /**
     * 追加数据--输出到文件
     * @param $data 二维数组
     * @return array
     * @create by hualu
     */
    private function _appendData()
    {
        //设置浏览器输出
        if ($this->_isOutput) {
            $this->_setHeader();
        }

        //打开文件句柄
        $this->_setFileop();
        //fwrite($this->_fileop, "\xEF\xBB\xBF" . 'sep=' . $this->_delimiter . $this->_ending);

        $line = '';
        $lineCount = 0;
        foreach ($this->_data as $item) {
            //分多次写入
            $line .= $this->_writeLine($item);
            $lineCount ++;
            if ($lineCount > $this->_limitCount) {
                fwrite($this->_fileop, $this->_getEncodingContent($line));
                $this->_resetFileop();
                $line = '';
                $lineCount = 0;
            }
        }
        fwrite($this->_fileop, $this->_getEncodingContent($line));
        $this->_data = null;
        $this->_closeFileop();
    }

    private function _setHeader()
    {
        if (!$this->_headerExists) {
            // 输出Excel表格到浏览器下载
            header('Content-Type: application/vnd.ms-excel;charset=' . $this->_encoding);
            header('Content-Disposition: attachment;filename="' . $this->_filename . '"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');
            // If you're serving to IE over SSL, then the following may be needed
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
            header('Pragma: public'); // HTTP/1.0
            $this->_headerExists = 1;
        }
    }

    /**
     * 单行数据处理
     * @param $line
     * @return string
     * @create by hualu
     */
    private function _writeLine($rowArr)
    {
        $element = '';
        if (!empty($rowArr) && is_array($rowArr)) {
            foreach ($rowArr as $v) {
                $line = str_replace($this->_enclosure, $this->_enclosure . $this->_enclosure, $v);
                //$line = str_replace($this->_delimiter, '，', $line);
                $element .= $this->_enclosure . $line . $this->_enclosure . $this->_delimiter;
            }
            $element .= $this->_ending;
        }
        return $element;
    }

    /**
     * 初始化文件名、文件路径
     * @create by hualu
     */
    private function _initFilePath()
    {
        //设置时间戳
        if (!$this->_timestamp) {
            $this->_timestamp = time();

            //文件名处理(默认)
            if (empty($this->_filename)) {
                $this->_filename = '数据导出_' . date('Y-m-d') . '_' . $this->_timestamp;
                //后缀统一
            } else {
                $lastPosition = strrpos($this->_filename, '.');
                $this->_filename =
                    $lastPosition ? substr($this->_filename, 0, $lastPosition) : $this->_filename;
            }
            $this->_filename = $this->_filename . "." . $this->_postFix;

            //文件路径处理
            $this->_path = rtrim($this->_path, '/') . '/';

            //第一次新写入
            $this->_isAppend = 0;
        } else {
            //之后的每一次均从文件尾写入
            $this->_isAppend = 1;
        }
    }

    /**
     * 参数校验
     */
    private function _paramsCheck($data)
    {
        if (!is_dir($this->_path)) {
            $this->_errorCode = self::ERROR_PATH;
            return false;
        }

        //校验头部
        if ($this->_head && !$this->_isValidHead($this->_head)) {
            return false;
        }

        //校验传入数组
        if (!$this->_isValidData($data)) {
            return false;
        }
        $this->_data = $data;

        return true;
    }

    /**
     * 列标题格式校验
     */
    private function _isValidHead($head)
    {
        //头部必须是一维数组
        if (!$this->_isOneDimension($head)) {
            $this->_errorCode = SELF::ERROR_HEADVALID;
            return false;
        }

        return true;
    }

    /**
     * 数据格式校验
     */
    private function _isValidData($data)
    {
        if (!is_array($data)) {
            $this->_errorCode = SELF::ERROR_DATAVALID;
            return false;
        }

        //数据必须为二维数组
        foreach ($data as $item) {
            if (!is_array($item)) {
                $this->_errorCode = SELF::ERROR_DATAVALID;
                return false;
            }

            foreach ($item as $v) {
                //检测是否为标量
                if (!is_scalar($v)) {
                    $this->_errorCode = SELF::ERROR_DATAVALID;
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * 判断某个数组是不是一维数组
     * @param $data
     * @return bool
     * @create by hualu
     */
    private function _isOneDimension($data)
    {
        if (!is_array($data)) {
            return false;
        }

        foreach ($data as $item) {
            if (!is_scalar($item)) {
                return false;
            }
        }
        return true;
    }

}
