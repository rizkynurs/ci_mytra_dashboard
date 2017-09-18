<?php 
	header('Content-Type: application/json');
	$host = "172.30.52.16";
	$port = "5432";
	$dbname = "mytra_dashboard";
	$user = "postgres";
	$password = "Mytr@123";
	$db = pg_connect( "host=$host port=$port dbname=$dbname user=$user password=$password" );

	$res1 = pg_query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-19' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res2 = pg_query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-06-26' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res3 = pg_query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-03' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res4 = pg_query("SELECT d.date, sum(sum(total)) over() as sum from (SELECT to_char(date_trunc('day', (date '2017-07-10' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join invoice i on d.date=to_char(date_trunc('day', i.invoice_date), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	//execute query
	$result1 = pg_fetch_assoc($res1);
	$result2 = pg_fetch_assoc($res2);
	$result3 = pg_fetch_assoc($res3);
	$result4 = pg_fetch_assoc($res4);

	$data[] = $result1;
	$data[] = $result2;
	$data[] = $result3;
	$data[] = $result4;

	//print_r($data);

	print json_encode($data);
?>