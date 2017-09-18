<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Dashboard extends CI_Controller {
	
	// Index login
	public function index() {
		$data = array(	'title'	=> 'Dashboard Mytra',
						'isi'	=> 'index');
		$this->load->view('layout/wrapper',$data);
	}

		public function event_chart()
		{
			$data = array('title' => 'Event Guard Charts', 
							'isi'=> 'pages/charts/event-guards');
			$this->load->view('layout_chart/wrapper',$data);
		}

}