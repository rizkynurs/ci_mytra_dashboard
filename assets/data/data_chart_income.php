<?php

	header('Content-Type: application/json');
	$host = "172.30.52.16";
	$port = "5432";
	$dbname = "mytra_dashboard";
	$user = "postgres";
	$password = "Mytr@123";
	$db = pg_connect( "host=$host port=$port dbname=$dbname user=$user password=$password" );

	$res = pg_query("select d.date, sum(total) from (select to_char(date_trunc('day', (date '2017-08-20' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') group by d.date order by d.date asc;");

	//execute query
	$result = pg_fetch_all($res);

	//loop through the returned data
	$data = array();
	foreach ($result as $row) {
		$data[] = $row;
	}

	//free memory associated with result
	//$result->close();

	//close connection
	//$db->close();

	//now print the data
	print json_encode($data);

?>