<?php if(! defined('BASEPATH')) exit('Akses langsung tidak diperbolehkan'); 

class Simple_login {
	// SET SUPER GLOBAL
	var $CI = NULL;
	public function __construct() {
		$this->CI =& get_instance();
	}
	// Fungsi login
	public function login($username, $password) {
		$this->CI->db->select('username','password');
		$this->CI->db->from('admin');
		$this->CI->db->where(array('username'=>$username,'password' => $password));
		$query = $this->CI->db->get();
		//$query = $this->CI->db->get_where('admin1',array('username'=>$username,'password' => $password));
		if($query->num_rows() == 1) {
			$this->CI->db->select('*');
	  		$this->CI->db->from('admin');
	  		$this->CI->db->where("username = '".$username."'");
			$row 	= $this->CI->db->get();
			$admin 	= $row->row();
			$id 	= $admin->id;
			$this->CI->session->set_userdata('username', $username);
			$this->CI->session->set_userdata('id_login', uniqid(rand()));
			$this->CI->session->set_userdata('id', $id);
			
			redirect(base_url('dashboard'));
		}else{
			$this->CI->session->set_flashdata('sukses','Oops... Username/password salah');
			redirect(base_url('login'));
		}
		return false;
	}
	// Proteksi halaman
	public function cek_login() {
		if($this->CI->session->userdata('username') == '') {
			$this->CI->session->set_flashdata('sukses','Anda belum login');
			redirect(base_url('login'));
		}
	}
	// Fungsi logout
	public function logout() {
		$this->CI->session->unset_userdata('username');
		$this->CI->session->unset_userdata('id_login');
		$this->CI->session->unset_userdata('id');
		$this->CI->session->set_flashdata('sukses','Anda berhasil logout');
		redirect(base_url('login'));
	}
}