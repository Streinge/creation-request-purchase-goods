<?php

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$checkedGoodsFilePath = '../files/checkedGoods.xlsx';
$checkedGoodsSpreadSheet = IOFactory::load($checkedGoodsFilePath);
$checkedGoodsSheet = $checkedGoodsSpreadSheet->getActiveSheet();

$dataCheckedGoods = [];

foreach ($checkedGoodsSheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();

    $rowData = [];
    $keys = ['ozm', 'name', 'unit', 'invalidBalance', 'purchaseVolume'];
    $i = 0;
    foreach ($cellIterator as $cell) {
        if ($i < count($keys)) {
            $rowData[$keys[$i]] = $cell->getValue();
            $i++;
        }
    }
    $dataCheckedGoods[] = $rowData;
}

$remantsWhFilePath = '../files/warehouseRemnants_09_01_2025.xlsx';
$remantsWhSpreadSheet = IOFactory::load($remantsWhFilePath);
$remantsWhSheet = $remantsWhSpreadSheet->getActiveSheet();

$dataRemantsWh = [];

foreach ($remantsWhSheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();

    $rowData = [];
    $keys = ['ozm', 'name', 'remant', 'unit'];
    $i = 0;
    foreach ($cellIterator as $cell) {
        if ($i === 0) {
            $ozm = (string) ($cell->getValue());
        }

        if ($i < count($keys)) {
            if ($i === 2) {
                $remant = (string) $cell->getValue();
                $remant = (int) str_replace('.', '', $remant);
                $rowData[$keys[$i]] = $remant;
                $i++;
                continue;
            }
            $rowData[$keys[$i]] = $cell->getValue();
            $i++;
        }
    }
    $dataRemantsWh[$ozm] = $rowData;
}



$purchaseRequest[] = ['ОЗМ', 'Наименование', 'Объем закупки', 'Единицы измерения', 'Остаток'];
foreach ($dataCheckedGoods as $checkedGood) {
    $ozm = (string) $checkedGood['ozm'];
    $purchaseVolume = $checkedGood['purchaseVolume'];
    $name = $checkedGood['name'];
    $unit = $checkedGood['unit'];
    if (isset($dataRemantsWh[$ozm]) && ((int) $dataRemantsWh[$ozm]['remant'] > (int) $checkedGood['invalidBalance'])) {
        continue;
    }

    if (isset($dataRemantsWh[$ozm])) {
        $name = $dataRemantsWh[$ozm]['name'] ?? '';
        $unit = $dataRemantsWh[$ozm]['unit'] ?? '';
    }

    $purchaseRequest[] = [
        'ozm' => $ozm,
        'name' => $name,
        'purchaseVolume' => $purchaseVolume,
        'unit' => $unit,
        'remantWh' => $dataRemantsWh[$ozm]['remant'] ?? '0'
    ];
}

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

foreach ($purchaseRequest as $rowIndex => $row) {
    foreach ($row as $colIndex => $value) {
        $sheet->fromArray($row, NULL, 'A' . ($rowIndex + 1));
    }
}


$writer = new Xlsx($spreadsheet);
$writer->save('../files/purchaseRequest_09_01_2025.xlsx');