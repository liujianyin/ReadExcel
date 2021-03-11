<?php
/**
 * Created by PhpStorm.
 * User: liujianyin.annabelle
 * Date: 2018/10/9
 * Time: 13:57
 */
//首先导入PHPExcel
set_time_limit(0);
header("Content-type:text/html;charset=utf-8");
require_once 'PHPExcel/Classes/PHPExcel.php';

$filePath = "20210303dc.xlsx";

//建立reader对象
$PHPReader = new PHPExcel_Reader_Excel2007();
if(!$PHPReader->canRead($filePath)){
    $PHPReader = new PHPExcel_Reader_Excel5();
    if(!$PHPReader->canRead($filePath)){
        echo 'no Excel';
        return ;
    }
}

//建立excel对象，此时你即可以通过excel对象读取文件，也可以通过它写入文件
$PHPExcel = $PHPReader->load($filePath);

/**读取excel文件中的第一个工作表*/
$currentSheet = $PHPExcel->getSheet(0);
/**取得最大的列号*/
$allColumn = $currentSheet->getHighestColumn();
/**取得一共有多少行*/
$allRow = $currentSheet->getHighestRow();
//file_put_contents('20210303dc.txt', "insert into  `app_ng_word_mst` (`word`,`ngworldversion`) VALUES", FILE_APPEND);
//循环读取每个单元格的内容。注意行从1开始，列从A开始
for($rowIndex=1;$rowIndex<=$allRow;$rowIndex++){
    for($colIndex='A';$colIndex<=$allColumn;$colIndex++){
        $addr = $colIndex.$rowIndex;
        $cell = $currentSheet->getCell($addr)->getValue();
//        if($cell instanceof PHPExcel_RichText)     //富文本转换字符串
//            $cell = $cell->__toString();
//        $cell = iconv('gb2312','utf-8',$cell);
//        echo iconv("GB2312","UTF-8",$cell);
//        echo $cell.setOutputEncoding('UTF-8'); ;
//        echo "/br";
//        echo $cell."\n";

        file_put_contents('20210303dc.txt', "'".$cell."',\n", FILE_APPEND);

    }
}

