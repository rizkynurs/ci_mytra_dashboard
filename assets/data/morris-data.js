$(function() {

    Morris.Area({
        element: 'morris-area-chart',
        data: [{
            Date: '2010 Q1',
            HG: 2666,
            EG: 0,
            
        }, {
            Date: '2010 Q2',
            HG: 5000,
            EG: 2500,
            
        }, {
            Date: '2010 Q3',
            HG: 4912,
            EG: 1969,
            
        }, {
            Date: '2010 Q4',
            HG: 3767,
            EG: 3597,
            
        }, {
            Date: '2011 Q1',
            HG: 6810,
            EG: 1914,
            
        }, {
            Date: '2011 Q2',
            HG: 5670,
            EG: 4293,
            
        }, {
            Date: '2011 Q3',
            HG: 4820,
            EG: 3795,
            
        }, {
            Date: '2011 Q4',
            HG: 15073,
            EG: 5967,
            
        }, {
            Date: '2012 Q1',
            HG: 10687,
            EG: 4460,
            
        }, {
            Date: '2012 Q2',
            HG: 8432,
            EG: 5713,
            
        }],
        xkey: 'Date',
        ykeys: ['HG', 'EG'],
        labels: ['Home Guards', 'Event Guards'],
        pointSize: 2,
        hideHover: 'auto',
        resize: true
    });

    Morris.Donut({
        element: 'morris-donut-chart',
        data: [{
            label: "Download Sales",
            value: 12
        }, {
            label: "In-Store Sales",
            value: 30
        }, {
            label: "Mail-Order Sales",
            value: 20
        }],
        resize: true
    });

    Morris.Bar({
        element: 'morris-bar-chart',
        data: [{
            y: '2006',
            a: 100,
            b: 90
        }, {
            y: '2007',
            a: 75,
            b: 65
        }, {
            y: '2008',
            a: 50,
            b: 40
        }, {
            y: '2009',
            a: 75,
            b: 65
        }, {
            y: '2010',
            a: 50,
            b: 40
        }, {
            y: '2011',
            a: 75,
            b: 65
        }, {
            y: '2012',
            a: 100,
            b: 90
        }],
        xkey: 'y',
        ykeys: ['a', 'b'],
        labels: ['Series A', 'Series B'],
        hideHover: 'auto',
        resize: true
    });
	
	Morris.Line({
		element: 'myfirstchart',  
		data: [{ 
			year: '2008', 
			value: 20 
		},{ 
			year: '2009', 
			value: 10 
		},{ 
			year: '2010', 
			value: 5 
		},{ 
			year: '2011', 
			value: 5 
		},{ 
			year: '2012', 
			value: 20 
		}],
		xkey: 'year',
		ykeys: ['value'],
		labels: ['Value']
});
    
});
