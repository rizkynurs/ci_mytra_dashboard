$(document).ready(function(){

	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_booking_daily_eg.php",
		type : "POST",
		//dataType : "json"
    	data : {pass_date: '2017-06-19'},
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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_booking_weekly_eg.php",
		type : "POST",
		success : function(data){
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
						label: "Event Guards Weekly Book",												
						backgroundColor: "rgba(10, 252, 35, 0.2)",
						borderColor: "rgba(10, 252, 35, 1)",
						pointHoverBackgroundColor: "rgba(10, 252, 35, 1)",
						pointHoverBorderColor: "rgba(10, 252, 35, 1)",
						borderWidth: 1,
						data: jumlah						
					}
				]
			};

			var ctx = $("#egbookweekly");

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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_booking_monthly_eg.php",
		type : "GET",
		success : function(data){
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
						label: "Event Guards Monthly Book",												
						backgroundColor: "rgba(25, 25, 25, 0.2)",
						borderColor: "rgba(25, 25, 25, 1)",
						pointHoverBackgroundColor: "rgba(25, 25, 25, 1)",
						pointHoverBorderColor: "rgba(25, 25, 25, 1)",
						borderWidth: 1,
						data: jumlah						
					}
				]
			};

			var ctx = $("#egbookmonthly");

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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_booking_yearly_eg.php",
		type : "GET",
		success : function(data){
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
						label: "Event Guards Yearly Book",												
						backgroundColor: "rgba(253, 102, 105, 0.2)",
						borderColor: "rgba(253, 102, 105, 1)",
						pointHoverBackgroundColor: "rgba(253, 102, 105, 1)",
						pointHoverBorderColor: "rgba(253, 102, 105, 1)",
						borderWidth: 1,
						data: jumlah						
					}
				]
			};

			var ctx = $("#egbookyearly");

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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_income_daily_eg.php",
		type : "GET",
		success : function(data){
			console.log(data);

			var inv_date = [];
			var jumlah = [];			

			for(var i in data) {
				inv_date.push(data[i].date);
				jumlah.push(data[i].sum);				
			}

			var chartdata = {
				labels: inv_date,
				datasets: [
					{
						label: 'Event Guards Daily Income',												
						backgroundColor: 'rgba(59, 89, 252, 0.2)',
						borderColor: 'rgba(59, 89, 252, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(59, 89, 252, 1)",
						pointHoverBorderColor: "rgba(59, 89, 252, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#egincomedaily");

			var LineGraph = new Chart(ctx, {
				type: 'bar',
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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_income_weekly_eg.php",
		type : "GET",
		success : function(data){
			console.log(data);

			var inv_date = [];
			var jumlah = [];			

			for(var i in data) {
				inv_date.push(data[i].date);
				jumlah.push(data[i].sum);				
			}

			var chartdata = {
				labels: inv_date,
				datasets: [
					{
						label: 'Event Guards Weekly Income',												
						backgroundColor: 'rgba(10, 252, 35, 0.2)',
						borderColor: 'rgba(10, 252, 35, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(10, 252, 35, 1)",
						pointHoverBorderColor: "rgba(10, 252, 35, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#egincomeweekly");

			var LineGraph = new Chart(ctx, {
				type: 'bar',
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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_income_monthly_eg.php",
		type : "GET",
		success : function(data){
			console.log(data);

			var inv_date = [];
			var jumlah = [];			

			for(var i in data) {
				inv_date.push(data[i].date);
				jumlah.push(data[i].sum);				
			}

			var chartdata = {
				labels: inv_date,
				datasets: [
					{
						label: 'Event Guards Monthly Income',												
						backgroundColor: 'rgba(25, 25, 25, 0.2)',
						borderColor: 'rgba(25, 25, 25, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(25, 25, 25, 1)",
						pointHoverBorderColor: "rgba(25, 25, 25, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#egincomemonthly");

			var LineGraph = new Chart(ctx, {
				type: 'bar',
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

$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/ci_mytra_dashboard/assets/data/data_chart_income_yearly_eg.php",
		type : "GET",
		success : function(data){
			console.log(data);

			var inv_date = [];
			var jumlah = [];			

			for(var i in data) {
				inv_date.push(data[i].date);
				jumlah.push(data[i].sum);				
			}

			var chartdata = {
				labels: inv_date,
				datasets: [
					{
						label: 'Event Guards Yearly Income',												
						backgroundColor: 'rgba(253, 102, 105, 0.2)',
						borderColor: 'rgba(253, 102, 105, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(253, 102, 105, 1)",
						pointHoverBorderColor: "rgba(253, 102, 105, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#egincomeyearly");

			var LineGraph = new Chart(ctx, {
				type: 'bar',
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
