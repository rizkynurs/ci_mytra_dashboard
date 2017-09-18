$(document).ready(function(){
	$.ajax({
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_booking_daily_hg",
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
						label: "Home Guards Daily Book",												
						backgroundColor: "rgba(59, 89, 252, 0.2)",
						borderColor: "rgba(59, 89, 252, 1",
						pointHoverBackgroundColor: "rgba(59, 89, 252, 1)",
						pointHoverBorderColor: "rgba(59, 89, 252, 1)",
						borderWidth: 1,
						data: jumlah							
					}
				]
			};

			var ctx = $("#hgbookdaily");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_booking_weekly_hg",
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
						label: "Home Guards Weekly Book",												
						backgroundColor: "rgba(10, 252, 35, 0.2)",
						borderColor: "rgba(10, 252, 35, 1)",
						pointHoverBackgroundColor: "rgba(10, 252, 35, 1)",
						pointHoverBorderColor: "rgba(10, 252, 35, 1)",
						borderWidth: 1,
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgbookweekly");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_booking_monthly_hg",
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
						label: "Home Guards Monthly Book",												
						backgroundColor: "rgba(25, 25, 25, 0.2)",
						borderColor: "rgba(25, 25, 25, 1)",
						pointHoverBackgroundColor: "rgba(25, 25, 25, 1)",
						pointHoverBorderColor: "rgba(25, 25, 25, 1)",
						borderWidth: 1,
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgbookmonthly");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_booking_yearly_hg",
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
						label: "Home Guards Yearly Book",												
						backgroundColor: "rgba(253, 102, 105, 0.2)",
						borderColor: "rgba(253, 102, 105, 1)",
						pointHoverBackgroundColor: "rgba(253, 102, 105, 1)",
						pointHoverBorderColor: "rgba(253, 102, 105, 1)",
						borderWidth: 1,
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgbookyearly");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_income_daily_hg",
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
						label: 'Home Guards Daily Income',												
						backgroundColor: 'rgba(59, 89, 252, 0.2)',
						borderColor: 'rgba(59, 89, 252, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(59, 89, 252, 1)",
						pointHoverBorderColor: "rgba(59, 89, 252, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgincomedaily");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_income_weekly_hg",
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
						label: 'Home Guards Weekly Income',												
						backgroundColor: 'rgba(10, 252, 35, 0.2)',
						borderColor: 'rgba(10, 252, 35, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(10, 252, 35, 1)",
						pointHoverBorderColor: "rgba(10, 252, 35, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgincomeweekly");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_income_monthly_hg",
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
						label: 'Home Guards Monthly Income',												
						backgroundColor: 'rgba(25, 25, 25, 0.2)',
						borderColor: 'rgba(25, 25, 25, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(25, 25, 25, 1)",
						pointHoverBorderColor: "rgba(25, 25, 25, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgincomemonthly");

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
		url : "http://localhost:8080/mytra-Dashboard/startbootstrap-sb-admin-2-gh-pages/data/data_chart_income_yearly_hg",
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
						label: 'Home Guards Yearly Income',												
						backgroundColor: 'rgba(253, 102, 105, 0.2)',
						borderColor: 'rgba(253, 102, 105, 1)',
						borderWidth: 1,
						pointHoverBackgroundColor: "rgba(253, 102, 105, 1)",
						pointHoverBorderColor: "rgba(253, 102, 105, 1)",
						data: jumlah						
					}
				]
			};

			var ctx = $("#hgincomeyearly");

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
