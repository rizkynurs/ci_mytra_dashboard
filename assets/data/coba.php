<!DOCTYPE html>
<html>
<head>
	<title></title>
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>
	<script type="text/javascript">
		
		$(document).ready(function()
		{
			$value = $("#tempat_taro_json").text();
			console.log($value);

			$.ajax({
				url : "event-guards.php",
				type : "POST",
				data: $value
				success : 
				function(data){
					console.log(data);

					var inv_date = [];
					var jumlah = [];			

					for(var i in data) {
						inv_date.push(data[i].date);
						jumlah.push(data[i].count);				
					}

					var chartdata = {
						labels: inv_date,
						datasets: [
							{
								label: "Event Guards Daily Book",												
								backgroundColor: "rgba(59, 89, 252, 0.2)",
								borderColor: "rgba(59, 89, 252, 1",
								pointHoverBackgroundColor: "rgba(59, 89, 252, 1)",
								pointHoverBorderColor: "rgba(59, 89, 252, 1)",
								borderWidth: 1,
								data: jumlah							
							}
						]
					};

					var ctx = $("#egbookdaily");

					var LineGraph = new Chart(ctx, {
						type: 'line',
						data: chartdata,
						options: {
							scales: {
								yAxes: [{
									ticks: {
										beginAtZero:true
									}
								}]
							}
						}
					});
				},
				error : function(data) {

				}
			});
		});
	</script>
</head>
<body>

<!-- 

	
-->
<div id="tempat_taro_json" style="display:none">
	<?php

    //header('Content-Type: application/json');
        $db = pg_connect( "host=172.30.52.16 port=5432 dbname=mytra_dashboard user=postgres password=Mytr@123" );
        
        $lala = $_REQUEST['coba'];

        $res = pg_query("SELECT d.date, count(i.booking_id) from (select to_char(date_trunc('day', (date '$lala' + offs)), 'yyyy-mm-dd') as date from generate_series(0,6,1) as offs) d left outer join booking i on d.date=to_char(date_trunc('day', i.booking_from), 'yyyy-mm-dd') where product_name = 'EG' group by d.date order by d.date asc;");

        //execute query
        $result = pg_fetch_all($res);

        //loop through the returned data
        $data = array();
        foreach ($result as $row) {
            $data[] = $row;
        }   

        print json_encode($data);

    ?>
</div>
</body>
</html>