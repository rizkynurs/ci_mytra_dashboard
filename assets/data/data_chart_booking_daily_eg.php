<?php

	header('Content-Type: application/json');
	$host = "172.30.52.16";
	$port = "5432";
	$dbname = "mytra_dashboard";
	$user = "postgres";
	$password = "Mytr@123";
	$db = pg_connect( "host=$host port=$port dbname=$dbname user=$user password=$password" );

	//include 'index.php';
	//echo $tfDate;
	$test = $_REQUEST['pass_date'];		
	
	$res = pg_query("SELECT d.date, count(i.booking_id) from (select to_char(date_trunc('day', (date '$test' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking i on d.date=to_char(date_trunc('day', i.booking_from), 'yyyy-mm-dd') where product_name = 'Event Guards' group by d.date order by d.date asc;");

	//execute query
	$result = pg_fetch_all($res);

	//loop through the returned data
	$data = array();
	foreach ($result as $row) {
		$data[] = $row;	
	}	

	print json_encode($data);

?>