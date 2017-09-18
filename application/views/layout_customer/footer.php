  <!-- jQuery -->
    <!--    <script src="../vendor/jquery/jquery.min.js"></script> -->
    <script src="<?php echo base_url(); ?>assets/vendor/jquery/jquery-1.12.4.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="<?php echo base_url(); ?>assets/vendor/bootstrap/js/bootstrap.min.js"></script>

    <!-- Metis Menu Plugin JavaScript -->
    <script src="<?php echo base_url(); ?>assets/vendor/metisMenu/metisMenu.min.js"></script>

    <!-- DataTables JavaScript -->
    <script src="<?php echo base_url(); ?>assets/vendor/datatables/js/jquery.dataTables.min.js"></script>
    <script src="<?php echo base_url(); ?>assets/vendor/datatables-plugins/dataTables.bootstrap.min.js"></script>
    <script src="<?php echo base_url(); ?>assets/vendor/datatables-responsive/dataTables.responsive.js"></script>
    <script src="<?php echo base_url(); ?>assets/vendor/datatables/js/dataTables.fixedColumns.min.js"></script>
    <!-- Custom Theme JavaScript -->
    <script src="<?php echo base_url(); ?>assets/dist/js/sb-admin-2.js"></script>
    <script src="<?php echo base_url(); ?>assets/dist/js/scroll.js"></script>
 <!-- Page-Level Demo Scripts - Tables - Use for reference -->
    <script>
    
    $(document).ready(function() {
    // Setup - add a text input to each footer cell
    $('#dataTables-example tfoot th').each( function (i) {
        var title = $(this).text();
        $(this).html( '<input type="text" placeholder="Search '+title+'" />' );
    } );
    
    $('#dataTables-example tbody').on( 'click', 'tr', function () {
        $(this).toggleClass('selected');
    } );

   var table = $('#dataTables-example').DataTable({
        "scrollX": true

   });
    
    // Apply the search
    table.columns().every( function () {
        var that = this;
 
        $( 'input', this.footer() ).on( 'keyup change', function () {
            if ( that.search() !== this.value ) {
                that
                    .search( this.value )
                    .draw();
            }
        } );
    } );
} );
    </script>

</body>

</html>