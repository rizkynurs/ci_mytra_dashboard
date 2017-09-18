<?php 
header('Content-Type: application/json');
	$host = "172.30.52.16";
	$port = "5432";
	$dbname = "mytra_dashboard";
	$user = "postgres";
	$password = "Mytr@123";
	$db = pg_connect( "host=$host port=$port dbname=$dbname user=$user password=$password" );

	$res1 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-01-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res2 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-02-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,27,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res3 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-03-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res4 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-04-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res5 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-05-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res6 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-06-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res7 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-07-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res8 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-08-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res9 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-09-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res10 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-10-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res11 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-11-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,29,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");

	$res12 = pg_query("SELECT d.date, sum(count(b.booking_id)) over() as count from (SELECT to_char(date_trunc('day', (date '2017-12-01' + offs)), 'yyyy-mm-dd') as date from generate_series(0,30,1) as offs) d left outer join booking b on d.date=to_char(date_trunc('day', b.booking_from), 'yyyy-mm-dd') where product_name = 'Home Guards' group by d.date order by d.date asc limit 1;");	

	//execute query
	$result1 = pg_fetch_assoc($res1);
	$result2 = pg_fetch_assoc($res2);
	$result3 = pg_fetch_assoc($res3);
	$result4 = pg_fetch_assoc($res4);
	$result5 = pg_fetch_assoc($res5);
	$result6 = pg_fetch_assoc($res6);
	$result7 = pg_fetch_assoc($res7);
	$result8 = pg_fetch_assoc($res8);
	$result9 = pg_fetch_assoc($res9);
	$result10 = pg_fetch_assoc($res10);
	$result11 = pg_fetch_assoc($res11);
	$result12 = pg_fetch_assoc($res12);


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

	//print_r($data);

	print json_encode($data);
?>