//Flot Line Chart

var d3 = [[1, 300], [2, 600], [3, 550], [4, 400], [5, 300]];
 
$(document).ready(function () {
    $.plot($("#eg-book-daily"), [d3]);
});

var d1 = [[1, 300], [2, 600], [3, 550], [4, 400], [5, 300]];
 
$(document).ready(function () {
    $.plot($("#eg-book-weekly"), [d1]);
});

var d2 = [[1, 500], [2, 100], [3, 950], [4, 250], [5, 800]];
 
$(document).ready(function () {
    $.plot($("#eg-book-monthly"), [d2]);
});

//Flot Line Chart
$(document).ready(function() {

    var offset = 0;
    plot();

    function plot() {
        var sin = [],
            cos = [];
        for (var i = 0; i < 12; i += 0.2) {
            sin.push([i, Math.sin(i + offset)]);
            cos.push([i, Math.cos(i + offset)]);
        }

        var options = {
            series: {
                lines: {
                    show: true
                },
                points: {
                    show: true
                }
            },
            grid: {
                hoverable: true //IMPORTANT! this is needed for tooltip to work
            },
            yaxis: {
                min: -1.2,
                max: 1.2
            },
            tooltip: true,
            tooltipOpts: {
                content: "'%s' of %x.1 is %y.4",
                shifts: {
                    x: -60,
                    y: 25
                }
            }
        };

        var plotObj = $.plot($("#eg-book-yearly"), [{
                data: sin,
                label: "sin(x)"
            }, {
                data: cos,
                label: "cos(x)"
            }],
            options);
    }
});

$(document).ready(function() {

    var barOptions1 = {
        series: {
            bars: {
                show: true,
                barWidth: 43200000
            }
        },
        xaxis: {
            mode: "time",
            timeformat: "%m/%d",
            minTickSize: [1, "day"]
        },
        grid: {
            hoverable: true
        },
        legend: {
            show: false
        },
        tooltip: true,
        tooltipOpts: {
            content: "x: %x, y: %y"
        }
    };
    var barData1 = {
        label: "bar",
        data: [
            [1354521600000, 1000],
            [1355040000000, 2000],
            [1355223600000, 3000],
            [1355306400000, 4000],
            [1355487300000, 5000],
            [1355571900000, 6000]
        ]
    };
    $.plot($("#eg-income-daily"), [barData1], barOptions1);

});


$(document).ready(function() {

    var barOptions1 = {
        series: {
            bars: {
                show: true,
                barWidth: 43200000
            }
        },
        xaxis: {
            mode: "time",
            timeformat: "%m/%d",
            minTickSize: [1, "day"]
        },
        grid: {
            hoverable: true
        },
        legend: {
            show: false
        },
        tooltip: true,
        tooltipOpts: {
            content: "x: %x, y: %y"
        }
    };
    var barData1 = {
        label: "bar",
        data: [
            [1354521600000, 1000],
            [1355040000000, 2000],
            [1355223600000, 3000],
            [1355306400000, 4000],
            [1355487300000, 5000],
            [1355571900000, 6000]
        ]
    };
    $.plot($("#eg-income-weekly"), [barData1], barOptions1);

});


$(document).ready(function() {

    var barOptions2 = {
        series: {
            bars: {
                show: true,
                barWidth: 43200000
            }
        },
        xaxis: {
            mode: "time",
            timeformat: "%m/%d",
            minTickSize: [1, "day"]
        },
        grid: {
            hoverable: true
        },
        legend: {
            show: false
        },
        tooltip: true,
        tooltipOpts: {
            content: "x: %x, y: %y"
        }
    };
    var barData2 = {
        label: "bar",
        data: [
            [1354521600000, 1000],
            [1355040000000, 2000],
            [1355223600000, 3000],
            [1355306400000, 4000],
            [1355487300000, 5000],
            [1355571900000, 6000]
        ]
    };
    $.plot($("#eg-income-monthly"), [barData2], barOptions2);

});

//Flot Bar Chart

$(document).ready(function() {

    var barOptions = {
        series: {
            bars: {
                show: true,
                barWidth: 86400000
            }
        },
        xaxis: {
            mode: "time",
            timeformat: "%m/%d",
            minTickSize: [7, "day"]
        },
        grid: {
            hoverable: true
        },
        legend: {
            show: false
        },
        tooltip: true,
        tooltipOpts: {
            content: "Tanggal: %x, Income: %y"
        }
    };
    var barData = {
        label: "bar",
        data: [
            [1501545600000, 1000],
            [1502150400000, 2000],
            [1502755200000, 3000],
            [1503360000000, 4000],
            [1503964800000, 10000],           
        ]
    };
    $.plot($("#eg-income-yearly"), [barData], barOptions);

});