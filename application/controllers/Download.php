<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Download extends CI_Controller {
	
	// Index login
	public function index() {
		$this->load->view('pages/download-number');
	}

		public function event_chart()
		{
			$data = array('title' => 'Event Guard Charts', 
							'isi'=> 'pages/charts/event-guards');
			$this->load->view('layout_chart/wrapper',$data);
		}

}