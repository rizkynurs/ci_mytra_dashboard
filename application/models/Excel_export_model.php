<?php
class Excel_export_model extends CI_Model
{
	function fetch_data_customer()
	{
	  $this->db->order_by("sys_user_id", "DESC");
	  $query = $this->db->get("customer");
	  return $query->result();
	}

	function fetch_data_booking()
	{
	  $this->db->order_by("booking_id", "DESC");
	  $query1 = $this->db->get("booking");
	  return $query1->result();
	}

	function fetch_data_invoice()
	{
	  $this->db->order_by("id", "DESC");
	  $query2 = $this->db->get("invoice");
	  return $query2->result();
	}

	function fetch_data_booking_home_guards()
	{
	  $this->db->select('*');
	  $this->db->from('booking');
	  $this->db->where("product_name = 'Home Guards'");
	  
	  $query3 = $this->db->get();
	  return $query3->result();
	}

	function fetch_data_booking_event_guards()
	{
	  $this->db->select('*');
	  $this->db->from('booking');
	  $this->db->where("product_name = 'Event Guards'");
	  
	  $query = $this->db->get();
	  return $query->result();
	}

	function fetch_data_invoice_home_guards()
	{
	  $this->db->select('*');
	  $this->db->from('invoice');
	  $this->db->where("product_name = 'Home Guards'");
	  
	  $query = $this->db->get();
	  return $query->result();
	}

	function fetch_data_invoice_event_guards()
	{
	  $this->db->select('*');
	  $this->db->from('invoice');
	  $this->db->where("product_name = 'Event Guards'");
	  
	  $query = $this->db->get();
	  return $query->result();
	}

}