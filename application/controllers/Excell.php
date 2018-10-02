<?php

defined('BASEPATH') OR exit('No direct script access allowed');

class Excell extends CI_Controller {

    function __construct()
    {
        parent::__construct();
        //load the excel library
        $this->load->library('excel');
        $this->load->model('kelas_model');
    }

    public function index()
    {
        $file = 'public/contoh_data_lengkap.xls';
        //read file from path
        $objPHPExcel = PHPExcel_IOFactory::load($file);

        //get only the Cell Collection
        $cell_collection = $objPHPExcel->getActiveSheet()->getCellCollection();

        //extract to a PHP readable array format
        foreach ($cell_collection as $cell)
        {
            $column = $objPHPExcel->getActiveSheet()->getCell($cell)->getColumn();
            $row = $objPHPExcel->getActiveSheet()->getCell($cell)->getRow();
            $data_value = $objPHPExcel->getActiveSheet()->getCell($cell)->getValue();

            //The header will/should be in row 1 only. of course, this can be modified to suit your need.
            if ($row == 1)
            {
                $header[$row][$column] = $data_value;
            }

            if (!empty($data_value) && $row != 1)
            {
                $arr_data[$row][$column] = $data_value;
            }
        }

        //send the data in an array format
        $data['header'] = $header;
        $data['values'] = $arr_data;
        
        $data_double = $this->_duplicate_data($data);
        // save data
        $this->_save($data['values'], $data_double);
    }

    public function writer()
    {
        //load our new PHPExcel library
        $this->load->library('excel');
        //activate worksheet number 1
        $this->excel->setActiveSheetIndex(0);
        //name the worksheet
        $this->excel->getActiveSheet()->setTitle('test worksheet');
        //set cell A1 content with some text
        $this->excel->getActiveSheet()->setCellValue('A1', 'This is just some text value');
        //change the font size
        $this->excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
        //make the font become bold
        $this->excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
        //merge cell A1 until D1
        $this->excel->getActiveSheet()->mergeCells('A1:D1');
        //set aligment to center for that merged cell (A1 to D1)
        $this->excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        $filename = 'just_some_random_name.xls'; //save our workbook as this file name
        header('Content-Type: application/vnd.ms-excel'); //mime type
        header('Content-Disposition: attachment;filename="' . $filename . '"'); //tell browser what's the file name
        header('Cache-Control: max-age=0'); //no cache
        //save it to Excel5 format (excel 2003 .XLS file), change this to 'Excel2007' (and adjust the filename extension, also the header mime type)
        //if you want to save it as .XLSX Excel 2007 format
        $objWriter = PHPExcel_IOFactory::createWriter($this->excel, 'Excel5');
        //force user to download the Excel file without writing it to server's HD
        $objWriter->save('php://output');
    }

    public function _save($save = '', $duplicate = '')
    {
        foreach ($duplicate as $key => $value)
        {
            unset($save[$value]);
        }
        
        // input data to database
        $this->kelas_model->save($data);
    }

    public function _duplicate_data($data = '')
    {
        if (!is_array($data))
        {
            return FALSE;
        }

        // search data duplicate
        $out = array();
        $unknown_data = array();
        $double['data_duplicate']['B'] = array();
        $double['data_duplicate']['D'] = array();
        $double['data_duplicate']['F'] = array();
        foreach ($data['values'] as $key => $value)
        {
            foreach ($data['values'] as $key2 => $value2)
            {
                if (!array_key_exists('B', $value) && !array_key_exists('D', $value) && !array_key_exists('F', $value))
                {
                    $unknown_data['unknown_data'][$key2] = $value;
                }

                if ($key != $key2)
                {
                    if (($value['B'] == $value2['B']))
                    {
                        $double['data_duplicate']['B'][] = $key2;
                    }

                    if (($value['D'] == $value2['D']))
                    {
                        $double['data_duplicate']['D'][] = $key2;
                    }

                    if (($value['F'] == $value2['F']))
                    {
                        $double['data_duplicate']['F'][] = $key2;
                    }
                }
            }
        }

        return array_keys(
                array_count_values(
                        array_merge($double['data_duplicate']['B'], $double['data_duplicate']['D'], $double['data_duplicate']['F'])
                )
        );
    }
}
