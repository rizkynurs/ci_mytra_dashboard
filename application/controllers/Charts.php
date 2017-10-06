<?php
 
if (!defined('BASEPATH')) exit('No direct script access allowed');
 
/**
*
* @author http://roytuts.com
*/
class Charts extends CI_Controller {
 
function __construct() {
	parent::__construct();
	$this->load->model('chart_models');
	$this->load->library('form_validation');
}
 
public function index() {
	//$x['data']=$this->chart_models->get_chart_invoice();
	//$x['data1']=$this->chart_model->data_chart_income_yearly_eg();
	$this->load->view('event-guards');
}
 
function get_chart_invoice() {
	if (isset($_POST['start']) AND isset($_POST['end'])) {
			$start_date = $_POST['start'];
			$end_date = $_POST['end'];
			$results = $this->chart_models->get_chart_invoice($start_date, $end_date);
				if ($results === NULL) {
					echo json_encode('No record found');
				} else {
					echo json_encode($results);
				}
		} else {
			echo json_encode('Date must be selected.');
		}
	}
 
}
 
/* End of file highchart.php */
/* Location: ./application/controllers/highchart.php */