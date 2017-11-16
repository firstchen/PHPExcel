<?php
/**
 * User: huangjianfan
 * Date: 2017/4/10
 * Time: 下午7:37
 */

namespace app\library;

class ExcelHelper
{
    public function __construct()
    {
        require_once APP_ROOT . '/library/PHPExcel.php';
    }

    public function read($filename, $readerType, $encode = 'utf-8')
    {
        $phpExcel = new \PHPExcel();
        $objReader = \PHPExcel_IOFactory::createReader($readerType);
        $objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($filename);
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $highestRow = $objWorksheet->getHighestRow();
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        $excelData = array();
        for ($row = 1; $row <= $highestRow; $row++) {
            for ($col = 0; $col < $highestColumnIndex; $col++) {
                $excelData[$row][] = (string)$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
            }
        }
        return $excelData;
    }

    public function push($data, $title, $name = 'Excel')
    {
        $objPHPExcel = new \PHPExcel();
        /*以下是一些设置 ，什么作者  标题啊之类的*/
        $objPHPExcel->getProperties()->setCreator("美柚")
            ->setLastModifiedBy("美柚")
            ->setTitle($title);
//            ->setSubject("数据EXCEL导出")
//            ->setDescription("备份数据")
//            ->setKeywords("excel")
//            ->setCategory("result file");
        foreach ($data as $row => $v) {
            foreach ($v as $col => $vv) {
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValueByColumnAndRow($col, $row);
            }
        }
        $objPHPExcel->getActiveSheet()->setTitle($title);
        $objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $name . '.xls"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
    }

    public function exportOtherSchedule($timeHeader, $scheduleList, $posNameMap)
    {
        $objPHPExcel = new \PHPExcel();
        /*以下是一些设置 ，什么作者  标题啊之类的*/
        $objPHPExcel->getProperties()->setCreator("美柚")
            ->setLastModifiedBy("美柚")
            ->setTitle('排期表');
        // 月份
        $rowIndex = 1;
        $monthCol = 0;
        $activeSheet = $objPHPExcel->setActiveSheetIndex(0);
        $activeSheet->setCellValueByColumnAndRow($monthCol++, $rowIndex, "");
        foreach ($timeHeader as $month => $days) {
            $activeSheet->setCellValueByColumnAndRow($monthCol, $rowIndex, $month)
                ->mergeCellsByColumnAndRow($monthCol, $rowIndex, $monthCol + count($days) - 1, $rowIndex);
            $monthCol += count($days);
        }

        // 星期
        $rowIndex++;
        $weekCol = 0;
        $activeSheet->setCellValueByColumnAndRow($weekCol++, $rowIndex, "");
        foreach ($timeHeader as $month => $days) {
            foreach ($days as $dd) {
                $activeSheet->setCellValueByColumnAndRow($weekCol++, $rowIndex, $dd[0]);
            }
        }

        // 日
        $rowIndex++;
        $dayCol = 0;
        $activeSheet->setCellValueByColumnAndRow($dayCol++, $rowIndex, "");
        foreach ($timeHeader as $month => $days) {
            foreach ($days as $dd) {
                $activeSheet->setCellValueByColumnAndRow($dayCol++, $rowIndex, $dd[1]);
            }
        }

        // 数据
        $rowIndex++;
        foreach ($scheduleList as $pos => $marks) {
            $col = 0;
            list($appId, $pageCode, $posCode, $ordinal) = explode('_', $pos);
            $txt = $posNameMap[$appId . '_' . $posCode];
            if ($ordinal) {
                $txt .= '-' . $ordinal;
            }
            $firstCol = $col++;
            $activeSheet->setCellValueByColumnAndRow($firstCol, $rowIndex, $txt);
            $activeSheet->getColumnDimensionByColumn($firstCol)->setWidth(35);
            foreach ($marks as $mark) {
                $curCol = $col++;
                $activeSheet->setCellValueByColumnAndRow($curCol, $rowIndex, $mark['txt']);
                if ($mark['txt']) {
                    $activeSheet->getStyleByColumnAndRow($curCol, $rowIndex)->getFont()
                        ->getColor()->setRGB(\PHPExcel_Style_Color::COLOR_DARKRED);

                    $activeSheet->getStyleByColumnAndRow($curCol,$rowIndex)->getFill()
                            ->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
                    $activeSheet->getStyleByColumnAndRow($curCol,$rowIndex)->getFill()
                            ->getStartColor()->setRGB(substr($mark['mark_color'],1));
                }
//
            }
            $rowIndex++;
        }

        $objPHPExcel->getActiveSheet()->setTitle('其它广告位排期表');
        $objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="排期表.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
        exit;
    }

    public function exportCircleSchedule($timeHeader, $scheduleList, $posList,$title)
    {
        $objPHPExcel = new \PHPExcel();
        /*以下是一些设置 ，什么作者  标题啊之类的*/
        $objPHPExcel->getProperties()->setCreator("美柚")
            ->setLastModifiedBy("美柚")
            ->setTitle($title);
        // 月份
        $rowIndex = 1;
        $monthCol = 0;
        $activeSheet = $objPHPExcel->setActiveSheetIndex(0);
        $activeSheet->setCellValueByColumnAndRow($monthCol, $rowIndex, "")
            ->mergeCellsByColumnAndRow($monthCol, $rowIndex, $monthCol + 3 - 1, $rowIndex);
        $monthCol += 3;
        foreach ($timeHeader as $month => $days) {
            $colSpan = count($days) * count($posList);
            $activeSheet->setCellValueByColumnAndRow($monthCol, $rowIndex, $month)
                ->mergeCellsByColumnAndRow($monthCol, $rowIndex, $monthCol + $colSpan - 1, $rowIndex);
            $monthCol += $colSpan;
        }

        // 星期
        $rowIndex++;
        $weekCol = 0;
        $activeSheet->setCellValueByColumnAndRow($weekCol++, $rowIndex, "");
        $activeSheet->setCellValueByColumnAndRow($weekCol++, $rowIndex, "");
        $activeSheet->setCellValueByColumnAndRow($weekCol++, $rowIndex, "");
        $colSpan = count($posList);
        foreach ($timeHeader as $month => $days) {
            foreach ($days as $dd) {
                $activeSheet->setCellValueByColumnAndRow($weekCol, $rowIndex, $dd[0])
                    ->mergeCellsByColumnAndRow($weekCol, $rowIndex, $weekCol + $colSpan - 1, $rowIndex);
                $weekCol += $colSpan;
            }
        }

        // 日
        $rowIndex++;
        $dayCol = 0;
        $activeSheet->setCellValueByColumnAndRow($dayCol++, $rowIndex, "");
        $activeSheet->setCellValueByColumnAndRow($dayCol++, $rowIndex, "");
        $activeSheet->setCellValueByColumnAndRow($dayCol++, $rowIndex, "");
        $colSpan = count($posList);
        foreach ($timeHeader as $month => $days) {
            foreach ($days as $dd) {
                $activeSheet->setCellValueByColumnAndRow($dayCol, $rowIndex, $dd[1])
                    ->mergeCellsByColumnAndRow($dayCol, $rowIndex, $dayCol + $colSpan - 1, $rowIndex);
                $dayCol += $colSpan;
            }
        }

        // 广告位
        $rowIndex++;
        $posCol = 0;
        $activeSheet->setCellValueByColumnAndRow($posCol++, $rowIndex, "频道");
        $activeSheet->setCellValueByColumnAndRow($posCol++, $rowIndex, "级别");
        $posHeadCol=$posCol++;
        $activeSheet->setCellValueByColumnAndRow($posHeadCol, $rowIndex, "位置");
        $activeSheet->getColumnDimensionByColumn($posHeadCol)->setWidth(18);
        foreach ($timeHeader as $month => $days) {
            foreach ($days as $dd) {
                foreach ($posList as $pos => $simpleName) {
                    $activeSheet->setCellValueByColumnAndRow($posCol++, $rowIndex, $simpleName);
                }
            }
        }

        $rowIndex++;
        foreach ($scheduleList as $categoryInfo) {
            foreach ($categoryInfo['forum_list'] as $k => $forumInfo) {
                $col = 0;
                $rowSpan = count($categoryInfo['forum_list']);
                if ($k == 0) {
                    $activeSheet->setCellValueByColumnAndRow($col, $rowIndex, $categoryInfo['name'])
                        ->mergeCellsByColumnAndRow($col, $rowIndex, $col, $rowIndex + $rowSpan - 1);
                }
                $col++;
                $activeSheet->setCellValueByColumnAndRow($col++, $rowIndex, $forumInfo['grade']);
                $activeSheet->setCellValueByColumnAndRow($col++, $rowIndex, $forumInfo['name']);
                foreach ($forumInfo['schedule'] as $mark) {
                    $curCol = $col++;
                    if(!is_array($mark)){
                        $mark['txt'] = '';
                    }
                    $activeSheet->setCellValueByColumnAndRow($curCol, $rowIndex, $mark['txt']);
                    if ($mark['txt']) {
                        $activeSheet->getStyleByColumnAndRow($curCol, $rowIndex)->getFont()
                            ->getColor()->setRGB(\PHPExcel_Style_Color::COLOR_DARKRED);
                        $activeSheet->getStyleByColumnAndRow($curCol,$rowIndex)->getFill()
                            ->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
                        $activeSheet->getStyleByColumnAndRow($curCol,$rowIndex)->getFill()
                            ->getStartColor()->setRGB(substr($mark['mark_color'],1));
                    }
                }
                $rowIndex++;
            }
        }

        $objPHPExcel->getActiveSheet()->setTitle($title);
        $objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$title.'.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
        exit;
    }
    
    public function exportExcelByRows($rows,$title) 
    {
        $objPHPExcel = new \PHPExcel();
        /*以下是一些设置 ，什么作者  标题啊之类的*/
        $objPHPExcel->getProperties()->setCreator("美柚")
            ->setLastModifiedBy("美柚")
            ->setTitle($title);
        $activeSheet = $objPHPExcel->setActiveSheetIndex(0);
        foreach($rows as $rowKey => $row){
            foreach($row as $key => $val){
                $activeSheet->setCellValueByColumnAndRow($key, $rowKey+1, $val);
            }
        }
        
        $objPHPExcel->getActiveSheet()->setTitle($title);
        $objPHPExcel->setActiveSheetIndex(0);
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$title.'.xlsx"');
        header('Cache-Control: max-age=0');
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
        exit;
    }
}