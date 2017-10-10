<?php
defined('BASEPATH') OR exit('No direct script access allowed');

/**
* 
*/
class Tables extends CI_Controller
{
	public function customer(){
		$data = array(	'title'	=> 'Customer Tables Page',
						'isi'	=> 'pages/tables/customer-tables');
		$this->load->view('layout_customer/wrapper',$data);
	}

	public function booking(){
		$data = array(	'title'	=> 'Booking Tables Page',
						'isi'	=> 'pages/tables/booking-tables');
		$this->load->view('layout_customer/wrapper',$data);
	}

	public function invoice()
	{
		$data = array(	'title'	=> 'Invoice Tables Page',
						'isi'	=> 'pages/tables/invoice-tables');
		$this->load->view('layout_customer/wrapper',$data);
	}
}
