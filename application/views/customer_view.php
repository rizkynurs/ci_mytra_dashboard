<!DOCTYPE html>
<html>
    <head> 
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Simple Serverside Datatable Codeigniter Bootstrap with ColVis</title>
    <link href="<?php echo base_url('assets/vendor/bootstrap/css/bootstrap.min.css')?>" rel="stylesheet">
    <link href="<?php echo base_url('assets/vendor/datatables/css/dataTables.bootstrap.min.css')?>" rel="stylesheet">
    <link href="<?php echo base_url('assets/vendor/datatables.colvis/dataTables.colVis.css')?>" rel="stylesheet">
    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    </head> 
<body>
    <div class="container">
        <h1 style="font-size:20pt">Simple Serverside Datatable Codeigniter Bootstrap with ColVis</h1>

        <h3>Customers Data</h3>
        <br>
        <div class="row">
            <div class="col-md-12">
                <div id="colvis"></div>
            </div>
        </div>
        <table id="table" class="table table-striped table-bordered" cellspacing="0" width="100%">
            <thead>
                <tr>
                    <th>No</th>
                    <th>User ID</th>
                    <th>Name</th>
                    <th>Username</th>
                    <th>Email</th>
                    <th>Phone</th>
                    <th>Mobile</th>
                </tr>
            </thead>
            <tbody>
            </tbody>

            <tfoot>
                <tr>
                    <th>No</th>
                    <th>User ID</th>
                    <th>Name</th>
                    <th>Username</th>
                    <th>Email</th>
                    <th>Phone</th>
                    <th>Mobile</th>
                </tr>
            </tfoot>
        </table>
    </div>

<script src="<?php echo base_url('assets/vendor/jquery/jquery-2.2.3.min.js')?>"></script>
<script src="<?php echo base_url('assets/vendor/bootstrap/js/bootstrap.min.js')?>"></script>
<script src="<?php echo base_url('assets/vendor/datatables/js/jquery.dataTables.min.js')?>"></script>
<script src="<?php echo base_url('assets/vendor/datatables/js/dataTables.bootstrap.min.js')?>"></script>
<script src="<?php echo base_url('assets/vendor/datatables.colvis/dataTables.colVis.js')?>"></script>


<script type="text/javascript">

var table;

$(document).ready(function() {

    //datatables
    table = $('#table').DataTable({ 

        "processing": true, //Feature control the processing indicator.
        "serverSide": true, //Feature control DataTables' server-side processing mode.
        "order": [], //Initial no order.
        // Load data for the table's content from an Ajax source
        "ajax": {
            "url": "<?php echo site_url('customer/customer_list')?>",
            "type": "POST"
        },

        //Set column definition initialisation properties.
        "columnDefs": [
        { 
            "targets": [ 0 ], //first column / numbering column
            "orderable": false, //set not orderable
        },
        ],

    });

    var colvis = new $.fn.dataTable.ColVis(table); //initial colvis
    $('#colvis').html(colvis.button()); //add colvis button to div with id="colvis"

});

</script>

</body>
</html>