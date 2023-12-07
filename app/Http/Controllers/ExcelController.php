<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\Storage;

class ExcelController extends Controller
{
    public function downloadExcel() {
        $sourcePath = resource_path('files/EFilingSalev1.xlsx');
        // Load the existing template
        $template = IOFactory::load($sourcePath);

        // Get the active sheet
        $sheet = $template->getActiveSheet();

        $columns = 5;
        $rows = 10;

        // array with numeric values
        for ($i = 0; $i < $rows; $i++) {
            $rowArray = [];
            for ($j = 0; $j < $columns; $j++) {
                $value = ($i + 1) * ($j + 1);
                
                // Use numerical indices directly
                $sheet->setCellValueByColumnAndRow($j + 1, $i + 1, $value);
            }
        }


        $templatePath = storage_path('app/public/excel-exports/EFilingSalev1.xlsx'); 
        
        // Save the modified template to the storage path
        $writer = new Xlsx($template);
        $writer->save($templatePath);


        // Return a download response
        return response()->download($templatePath, 'EFilingSalev1.xlsx')->deleteFileAfterSend(true);
    }
}
