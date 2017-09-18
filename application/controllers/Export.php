<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Export extends CI_Controller {
 
   function customer()
   {
      $this->load->model("excel_export_model");
      $this->load->library("excel");
      $customers = new PHPExcel();

      $customers->setActiveSheetIndex(0);

      /*
      $table_columns = array("Name", "Address", "Gender", "Designation", "Age");

      $column = 0;

      foreach($table_columns as $field)
      {
       $object->getActiveSheet()->setCellValueByColumnAndRow($column, 1, $field);
       $column++;
      }
      
      */

      $customers->getActiveSheet()
          ->setCellValue('A1' , 'Data Customer')
          ->setCellValue('A3' , 'Name')
          ->setCellValue('B3' , 'Email')
          ->setCellValue('C3' , 'Phone Number')
          ->setCellValue('D3' , 'Mobile Phone Number');

      $customers->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal('center');
      $customers->getActiveSheet()->getStyle('A3:D3')->getAlignment()->setHorizontal('center');
      $customers->getActiveSheet()->setAutoFilter('A3:D3');
      $customers->getActiveSheet()->freezePane('A4');

      $customer_data = $this->excel_export_model->fetch_data_customer();

      //Set align
      
      $row = 4;

      foreach($customer_data as $result)
      {
       $customers->getActiveSheet()
              ->setCellValue('A'.$row , $result->name)
              ->setCellValue('B'.$row , $result->email)
              ->setCellValue('C'.$row , $result->phone)
              ->setCellValue('D'.$row , $result->mobile);
              $row++;
      }

      for($col = 'A'; $col !== 'E'; $col++) {
      $customers->getActiveSheet()
          ->getColumnDimension($col)
          ->setAutoSize(true);
      }
      

      //merging tittle (A1)
      $customers->getActiveSheet()->mergeCells('A1:D1');

      //style
      $customers->getActiveSheet()->getStyle('A1')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 20,
                  'name' => 'Tahoma' 
                  )
              )
          );

      $customers->getActiveSheet()->getStyle('A3:D3')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 12,
                  'name' => 'Tahoma' 
                  ),
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      //border for data
      $customers->getActiveSheet()->getStyle('A4:D'.($row-1))->applyFromArray(
          array(
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      

      //redirect to browser 
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment; filename="customer.xlsx"');
      header('Cache-Control: max-age=0');

      $file = PHPExcel_IOFactory::createWriter($customers,'Excel2007');

      $file->save('php://output');
   }


   function invoice()
   {
      $this->load->model("excel_export_model");
      $this->load->library("excel");
      $invoices = new PHPExcel();

      $invoices->setActiveSheetIndex(0);

      /*
      $table_columns = array("Name", "Address", "Gender", "Designation", "Age");

      $column = 0;

      foreach($table_columns as $field)
      {
       $object->getActiveSheet()->setCellValueByColumnAndRow($column, 1, $field);
       $column++;
      }
      
      */

      $invoices->getActiveSheet()
          ->setCellValue('A1' , 'Data Invoices')
          ->setCellValue('A3' , 'Number')
          ->setCellValue('B3' , 'Invoice Date')
          ->setCellValue('C3' , 'Invoice Due Date')
          ->setCellValue('D3' , 'Currency Code')
          ->setCellValue('E3' , 'Total')
          ->setCellValue('F3' , 'Booking Number')
          ->setCellValue('G3' , 'Product Name')
          ->setCellValue('H3' , 'Status')
          ->setCellValue('I3' , 'Customer Name')
          ->setCellValue('J3' , 'Customer Mobile')
          ->setCellValue('K3' , 'Customer Email');

      $invoices->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal('center');
      $invoices->getActiveSheet()->getStyle('A3:K3')->getAlignment()->setHorizontal('center');
      $invoices->getActiveSheet()->setAutoFilter('A3:K3');
      $invoices->getActiveSheet()->freezePane('A4');

      $invoice_data = $this->excel_export_model->fetch_data_invoice();

      //Set align
      
      $row = 4;

      foreach($invoice_data as $result)
      {
       $invoices->getActiveSheet()
              ->setCellValue('A'.$row , $result->number)
              ->setCellValue('B'.$row , $result->invoice_date)
              ->setCellValue('C'.$row , $result->invoice_duedate)
              ->setCellValue('D'.$row , $result->currency_code)
              ->setCellValue('E'.$row , $result->total)
              ->setCellValue('F'.$row , $result->booking_number)
              ->setCellValue('G'.$row , $result->product_name)
              ->setCellValue('H'.$row , $result->status)
              ->setCellValue('I'.$row , $result->customer_name)
              ->setCellValue('J'.$row , $result->customer_mobile)
              ->setCellValue('K'.$row , $result->customer_email);
              $row++;
      }

      for($col = 'A'; $col !== 'L'; $col++) {
      $invoices->getActiveSheet()
          ->getColumnDimension($col)
          ->setAutoSize(true);
      }
      

      //merging tittle (A1)
      $invoices->getActiveSheet()->mergeCells('A1:K1');

      //style
      $invoices->getActiveSheet()->getStyle('A1')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 20,
                  'name' => 'Tahoma' 
                  )
              )
          );

      $invoices->getActiveSheet()->getStyle('A3:K3')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 12,
                  'name' => 'Tahoma' 
                  ),
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      //border for data
      $invoices->getActiveSheet()->getStyle('A4:K'.($row-1))->applyFromArray(
          array(
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      

      //redirect to browser 
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment; filename="invoice.xlsx"');
      header('Cache-Control: max-age=0');

      $file = PHPExcel_IOFactory::createWriter($invoices,'Excel2007');

      $file->save('php://output');
     }
   
   function booking()
   {
      $this->load->model("excel_export_model");
      $this->load->library("excel");
      $bookings = new PHPExcel();

      $bookings->setActiveSheetIndex(0);

      /*
      $table_columns = array("Name", "Address", "Gender", "Designation", "Age");

      $column = 0;

      foreach($table_columns as $field)
      {
       $object->getActiveSheet()->setCellValueByColumnAndRow($column, 1, $field);
       $column++;
      }
      
      */

      $bookings->getActiveSheet()
        ->setCellValue('A1' , 'Data Booking') //title
        ->setCellValue('A3' , 'ID Booking')
        ->setCellValue('B3' , 'Address')
        ->setCellValue('C3' , 'PIC')
        ->setCellValue('D3' , 'Product Name')
        ->setCellValue('E3' , 'Booking From')
        ->setCellValue('F3' , 'Booking Until')
        ->setCellValue('G3' , 'Quantity')
        ->setCellValue('H3' , 'Price')
        ->setCellValue('I3' , 'Location')
        ->setCellValue('J3' , 'Status');

      $bookings->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal('center');
      $bookings->getActiveSheet()->getStyle('A3:J3')->getAlignment()->setHorizontal('center');
      $bookings->getActiveSheet()->setAutoFilter('A3:J3');
      $bookings->getActiveSheet()->freezePane('A4');

      $booking_data = $this->excel_export_model->fetch_data_booking();

      //Set align
      
      $row = 4;

      foreach($booking_data as $result)
      {
       $bookings->getActiveSheet()
            ->setCellValue('A'.$row , $result->booking_id)
            ->setCellValue('B'.$row , $result->address)
            ->setCellValue('C'.$row , $result->pic)
            ->setCellValue('D'.$row , $result->product_name)
            ->setCellValue('E'.$row , $result->booking_from)
            ->setCellValue('F'.$row , $result->booking_until)
            ->setCellValue('G'.$row , $result->qty)
            ->setCellValue('H'.$row , $result->price)
            ->setCellValue('I'.$row , $result->geolocation)
            ->setCellValue('J'.$row , $result->status);
            $row++;
      }

      for($col = 'A'; $col !== 'K'; $col++) {
      $bookings->getActiveSheet()
          ->getColumnDimension($col)
          ->setAutoSize(true);
      }
      

      //merging tittle (A1)
      $bookings->getActiveSheet()->mergeCells('A1:J1');

      //style
      $bookings->getActiveSheet()->getStyle('A1')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 20,
                  'name' => 'Tahoma' 
                  )
              )
          );

      $bookings->getActiveSheet()->getStyle('A3:J3')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 12,
                  'name' => 'Tahoma' 
                  ),
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      //border for data
      $bookings->getActiveSheet()->getStyle('A4:J'.($row-1))->applyFromArray(
          array(
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      

      //redirect to browser 
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment; filename="booking.xlsx"');
      header('Cache-Control: max-age=0');

      $file = PHPExcel_IOFactory::createWriter($bookings,'Excel2007');

      $file->save('php://output');
     }

     function booking_home_guards()
    {
      $this->load->model("excel_export_model");
      $this->load->library("excel");
      $bookinghg = new PHPExcel();

      $bookinghg->setActiveSheetIndex(0);

      /*
      $table_columns = array("Name", "Address", "Gender", "Designation", "Age");

      $column = 0;

      foreach($table_columns as $field)
      {
       $object->getActiveSheet()->setCellValueByColumnAndRow($column, 1, $field);
       $column++;
      }
      
      */

      $bookinghg->getActiveSheet()
        ->setCellValue('A1' , 'Data Booking') //title
        ->setCellValue('A3' , 'ID Booking')
        ->setCellValue('B3' , 'Address')
        ->setCellValue('C3' , 'PIC')
        ->setCellValue('D3' , 'Product Name')
        ->setCellValue('E3' , 'Booking From')
        ->setCellValue('F3' , 'Booking Until')
        ->setCellValue('G3' , 'Quantity')
        ->setCellValue('H3' , 'Price')
        ->setCellValue('I3' , 'Location')
        ->setCellValue('J3' , 'Status');

      $bookinghg->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal('center');
      $bookinghg->getActiveSheet()->getStyle('A3:J3')->getAlignment()->setHorizontal('center');
      $bookinghg->getActiveSheet()->setAutoFilter('A3:J3');
      $bookinghg->getActiveSheet()->freezePane('A4');

      $bookinghg_data = $this->excel_export_model->fetch_data_booking_home_guards();

      //Set align
      
      $row = 4;

      foreach($bookinghg_data as $result)
      {
       $bookinghg->getActiveSheet()
            ->setCellValue('A'.$row , $result->booking_id)
            ->setCellValue('B'.$row , $result->address)
            ->setCellValue('C'.$row , $result->pic)
            ->setCellValue('D'.$row , $result->product_name)
            ->setCellValue('E'.$row , $result->booking_from)
            ->setCellValue('F'.$row , $result->booking_until)
            ->setCellValue('G'.$row , $result->qty)
            ->setCellValue('H'.$row , $result->price)
            ->setCellValue('I'.$row , $result->geolocation)
            ->setCellValue('J'.$row , $result->status);
            $row++;
      }

      for($col = 'A'; $col !== 'K'; $col++) {
      $bookinghg->getActiveSheet()
          ->getColumnDimension($col)
          ->setAutoSize(true);
      }
      

      //merging tittle (A1)
      $bookinghg->getActiveSheet()->mergeCells('A1:J1');

      //style
      $bookinghg->getActiveSheet()->getStyle('A1')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 20,
                  'name' => 'Tahoma' 
                  )
              )
          );

      $bookinghg->getActiveSheet()->getStyle('A3:J3')->applyFromArray(
          array(
              'font' => array(
                  'bold' => true,
                  'size' => 12,
                  'name' => 'Tahoma' 
                  ),
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      //border for data
      $bookinghg->getActiveSheet()->getStyle('A4:J'.($row-1))->applyFromArray(
          array(
              'borders' => array(
                  'allborders' => array(
                      'style' => PHPExcel_Style_Border::BORDER_THIN
                      )
                  )
              )
          );

      

      //redirect to browser 
      header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      header('Content-Disposition: attachment; filename="booking home guard.xlsx"');
      header('Cache-Control: max-age=0');

      $file = PHPExcel_IOFactory::createWriter($bookinghg,'Excel2007');

      $file->save('php://output');
     }
}