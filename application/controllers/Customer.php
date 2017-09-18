<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Customer extends CI_Controller {

	public function __construct()
	{
		parent::__construct();
		$this->load->model('customer_model','customer');
	}

	public function index()
	{
		$this->load->helper('url');
		$this->load->view('customer_view');
	}

	public function customer_list()
	{
		$list = $this->customer->get_datatables();
		$data = array();
		$no = $_POST['start'];
		foreach ($list as $customer) {
			$no++;
			$row = array();
			$row[] = $no;
			$row[] = $customer->sys_user_id;
			$row[] = $customer->name;
			$row[] = $customer->username;
			$row[] = $customer->email;
			$row[] = $customer->phone;
			$row[] = $customer->mobile;

			$data[] = $row;
		}

		$output = array(
						"draw" => $_POST['draw'],
						"recordsTotal" => $this->customer->count_all(),
						"recordsFiltered" => $this->customer->count_filtered(),
						"data" => $data,
				);
		//output to json format
		echo json_encode($output);
	}

}
