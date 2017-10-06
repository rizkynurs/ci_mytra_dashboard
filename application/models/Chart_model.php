<?php
defined('BASEPATH') OR exit('No direct script access allowed');

/**
* 
*/
class Chart_model extends CI_Model
{

	function data_chart_booking()
	{
		$query=$this->db->query("SELECT d.date, count(i.id) from (select to_char(date_trunc('day', (date '2017-07-06' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc");

		if ($query->num_rows() > 0){
			foreach ($query->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}
	}

	function data_chart_booking_daily_eg()
	{
		$query= $this->db->query("SELECT d.date, count(i.booking_id) from (select to_char(date_trunc('day', (date '$test' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking i on d.date=to_char(date_trunc('day', i.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc;");
		
		if ($query->num_rows() > 0) {
			foreach ($query->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}
	}

	function data_chart_booking_daily_hg()
	{
		$query=$this->db->query("SELECT d.date, count(i.booking_id) from (select to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking i on d.date=to_char(date_trunc('day', i.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc;");

		if ($query->num_rows() > 0) {
			foreach ($query->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}
	}

	function data_chart_booking_monthly()
	{
		$res = $this->db->query("SELECT to_char(booking_from,'Month') as month, sum(count(booking_id)) over() as count from booking WHERE status='PAD' group by 1 ORDER BY month DESC;");

		if ($res->num_rows() > 0){
			foreach ($res->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}
		return NULL;

	}

	function data_chart_booking_monthly_hg()
	{
		$query1=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query1=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-02-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,27,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query2=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-03-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query3=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-04-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query4=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-05-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query5=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query6=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query7=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query8=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-08-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query9=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-09-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query10=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-10-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query11=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-11-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$query12=$this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (select to_char(date_trunc('day', (date '2017-12-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$result1 = pg_fetch_assoc($this->$query1);
		$result2 = pg_fetch_assoc($this->$query2);
		$result3 = pg_fetch_assoc($this->$query3);
		$result4 = pg_fetch_assoc($this->$query4);
		$result5 = pg_fetch_assoc($this->$query5);
		$result6 = pg_fetch_assoc($this->$query6);
		$result7 = pg_fetch_assoc($this->$query7);
		$result8 = pg_fetch_assoc($this->$query8);
		$result9 = pg_fetch_assoc($this->$query9);
		$result10 = pg_fetch_assoc($this->$query10);
		$result11 = pg_fetch_assoc($this->$query11);
		$result12 = pg_fetch_assoc($this->$query12);


		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
		$data[] = $result5;
		$data[] = $result6;
		$data[] = $result7;
		$data[] = $result8;
		$data[] = $result9;
		$data[] = $result10;
		$data[] = $result11;
		$data[] = $result12;
		
		return $data;
	}

	function data_chart_booking_monthly_eg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-02-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,27,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-03-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-04-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res5 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-05-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res6 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res7 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res8 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-08-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res9 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-09-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res10 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-10-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res11 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-11-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res12 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-12-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");	

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);
		$result5 = pg_fetch_assoc($this->$res5);
		$result6 = pg_fetch_assoc($this->$res6);
		$result7 = pg_fetch_assoc($this->$res7);
		$result8 = pg_fetch_assoc($this->$res8);
		$result9 = pg_fetch_assoc($this->$res9);
		$result10 = pg_fetch_assoc($this->$res10);
		$result11 = pg_fetch_assoc($this->$res11);
		$result12 = pg_fetch_assoc($this->$res12);


		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
		$data[] = $result5;
		$data[] = $result6;
		$data[] = $result7;
		$data[] = $result8;
		$data[] = $result9;
		$data[] = $result10;
		$data[] = $result11;
		$data[] = $result12;

		return $data;
	}

	function data_chart_booking_weekly()
	{
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-26' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
	
		return $data;
	}

	function data_chart_booking_weekly_eg()
	{
			
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-26' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
	
		return $data;
	}

	function data_chart_booking_weekly_hg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-26' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;
	}

	function data_chart_booking_yearly()
	{
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2018-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2019-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2020-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,366,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;	
	}

	function data_chart_booking_yearly_eg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2018-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2019-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2020-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,366,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;
	}

	function data_chart_booking_yearly_hg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2018-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2019-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,365,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2020-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,366,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;
	}

	function data_chart_income()
	{
		$res = $this->db->query("SELECT d.date, sum(total) from (select to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,31,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc;");

		if ($res->num_rows() > 0) {
			foreach ($res->result() as $row) {
				$data[]=$row;
			}			

			return $data;
		}
	}

	function data_chart_income_daily_eg()
	{
		$res = $this->db->query("SELECT d.date, sum(total) from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc;");
		
		if ($res->num_rows() > 0) {
			foreach ($res->result() as $ow) {
				$data[]=$row;	
			}

			return $data;
		}
	}

	function data_chart_income_daily_hg()
	{
		$res = $this->db->query("SELECT d.date, sum(total) from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc;");
		
		if ($res->num_rows() > 0) {
			foreach ($query->result() as $ow) {
				$data[]=$row;	
			}

			return $data;
		}
	}

	function data_chart_income_monthly()
	{
		$res1 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-02-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,27,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-03-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-04-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res5 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-05-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res6 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res7 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res8 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-08-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res9 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-09-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res10 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-10-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res11 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-11-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res12 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-12-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");	
	
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);
		$result5 = pg_fetch_assoc($this->$res5);
		$result6 = pg_fetch_assoc($this->$res6);
		$result7 = pg_fetch_assoc($this->$res7);
		$result8 = pg_fetch_assoc($this->$res8);
		$result9 = pg_fetch_assoc($this->$res9);
		$result10 = pg_fetch_assoc($this->$res10);
		$result11 = pg_fetch_assoc($this->$res11);
		$result12 = pg_fetch_assoc($this->$res12);


		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
		$data[] = $result5;
		$data[] = $result6;
		$data[] = $result7;
		$data[] = $result8;
		$data[] = $result9;
		$data[] = $result10;
		$data[] = $result11;
		$data[] = $result12;
		
		return $data;
	}

	function data_chart_income_monthly_eg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-02-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,27,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-03-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-04-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res5 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-05-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res6 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res7 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res8 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-08-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res9 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-09-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res10 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-10-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res11 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-11-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res12 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-12-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");	

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);
		$result5 = pg_fetch_assoc($this->$res5);
		$result6 = pg_fetch_assoc($this->$res6);
		$result7 = pg_fetch_assoc($this->$res7);
		$result8 = pg_fetch_assoc($this->$res8);
		$result9 = pg_fetch_assoc($this->$res9);
		$result10 = pg_fetch_assoc($this->$res10);
		$result11 = pg_fetch_assoc($this->$res11);
		$result12 = pg_fetch_assoc($this->$res12);


		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
		$data[] = $result5;
		$data[] = $result6;
		$data[] = $result7;
		$data[] = $result8;
		$data[] = $result9;
		$data[] = $result10;
		$data[] = $result11;
		$data[] = $result12;

		return $data;
	}

	function data_chart_income_monthly_hg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-02-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,27,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-03-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-04-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res5 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-05-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res6 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res7 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res8 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-08-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res9 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-09-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res10 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-10-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res11 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-11-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res12 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-12-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join invoice b on d.date=to_char(date_trunc('day', b.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");	

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);
		$result5 = pg_fetch_assoc($this->$res5);
		$result6 = pg_fetch_assoc($this->$res6);
		$result7 = pg_fetch_assoc($this->$res7);
		$result8 = pg_fetch_assoc($this->$res8);
		$result9 = pg_fetch_assoc($this->$res9);
		$result10 = pg_fetch_assoc($this->$res10);
		$result11 = pg_fetch_assoc($this->$res11);
		$result12 = pg_fetch_assoc($this->$res12);


		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;
		$data[] = $result5;
		$data[] = $result6;
		$data[] = $result7;
		$data[] = $result8;
		$data[] = $result9;
		$data[] = $result10;
		$data[] = $result11;
		$data[] = $result12;

		return $data;
	}

	function data_chart_income_weekly()
	{
		$res1 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-08-20' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-08-27' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-09-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (select to_char(date_trunc('day', (date '2017-09-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;
	}

	function data_chart_income_weekly_eg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-26' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;
	}

	function data_chart_income_weekly_hg()
	{
		$res1 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res2 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-26' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res3 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		$res4 = $this->db->query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

		//execute query
		$result1 = pg_fetch_assoc($this->$res1);
		$result2 = pg_fetch_assoc($this->$res2);
		$result3 = pg_fetch_assoc($this->$res3);
		$result4 = pg_fetch_assoc($this->$res4);

		$data[] = $result1;
		$data[] = $result2;
		$data[] = $result3;
		$data[] = $result4;

		return $data;
	}

	function data_chart_income_yearly()
	{
		$res = $this->db->query("SELECT to_char(invoice_date,'Month') as month, sum(total) as total from invoice WHERE status='PAD' group by 1 ORDER BY month DESC;");

		if ($res->num_rows() > 0){
			foreach ($res->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}
		return NULL;

		
	}

	function data_chart_income_yearly_eg()
	{
		$res = $this->db->query("SELECT to_char(invoice_date,'Month') as month, sum(total) as total from invoice WHERE status='PAD' AND product_name='Event Guards' group by 1 ORDER BY month DESC;");

		if ($res->num_rows() > 0){
			foreach ($res->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}


		else{
			echo "No Data Found";
		}
	}

	function data_chart_income_yearly_hg()
	{
		$res = $this->db->query("SELECT to_char(invoice_date,'Month') as month, sum(total) as total from invoice WHERE status='PAD' AND product_name='Home Guards' group by 1 ORDER BY month DESC;");

		if ($res->num_rows() > 0){
			foreach ($res->result() as $row) {
				$data[] = $row;
			}
			return $data;
		}

		else{
			echo "No Data Found";
	}
}
}