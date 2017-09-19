<div id="page-wrapper">
            <div class="row">
                <div class="col-lg-12">
                    <h1 class="page-header">Invoice Tables
                    <!-- <div class="pull-right">
                        <div class="btn-group">
                        <button type="button" class="btn btn-default btn-xs dropdown-toggle" data-toggle="dropdown">
                        Actions
                        <span class="caret"></span>
                        </button>
                             <ul class="dropdown-menu pull-right" role="menu">
                                <li><a href="#">Download to EXCEL</a></li>
                                <li><a href="#">Another action</a></li>
                                <li><a href="pdf_tables.php">Download to PDF</a></li>
                                <li class="divider"></li>
                                    <li><a href="#">Separated link</a>
                                </li>
                            </ul>
                        </div>
                    </div> -->
                    </h1>

                </div>
                <!-- /.col-lg-12 -->
            </div>
            <!-- /.row -->
            <div class="row">
                <div class="col-lg-12">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            Invoice's Tables
                            <div class="pull-right">
                            <div class="btn-group">
                                    <button type="button" class="btn btn-default btn-xs dropdown-toggle" data-toggle="dropdown">
                                        Actions
                                        <span class="caret"></span>
                                    </button>
                                    <ul class="dropdown-menu pull-right" role="menu">
                                        <li><a href="<?=base_url('export/invoice'); ?>">Download to EXCEL</a>
                                        </li>
                                        <li><a href="<?=base_url('export/invoice_event_guards'); ?>">Download Event Guard Invoices EXCEL</a>
                                        </li>
                                        <li><a href="<?=base_url('export/invoice_home_guards'); ?>">Download Home Guard Invoices EXCEL</a>
                                        </li>
                                        <li class="divider"></li>
                                        <li><a href="#">Separated link</a>
                                        </li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                        <!-- /.panel-heading -->
                        <div class="panel-body">
                            <table width="100%" class="table table-striped table-bordered table-hover" id="dataTables-example">
                            <!-- <table width="100%" class="table table-striped table-bordered table-hover" id="dataTables-example"> -->
                                <thead>
                                    <tr>
                                        <th>Number</th>
                                        <th>Invoice Date</th>
                                        <th>Invoice Due Date</th>
                                        <th>Currency Code</th>
                                        <th>Total</th>
                                        <th>Booking Number</th>
                                        <th>Product Name</th>
                                        <th>Status</th>
                                        <th>Customer Name</th>
                                        <th>Customer Mobile</th>
                                        <th>Customer Email</th>
                                    </tr>
                                </thead>    
                                <tfoot>
                                    <tr>
                                        <th>Number</th>
                                        <th>Invoice Date</th>
                                        <th>Invoice Due Date</th>
                                        <th>Currency Code</th>
                                        <th>Total</th>
                                        <th>Booking Number</th>
                                        <th>Product Name</th>
                                        <th>Status</th>
                                        <th>Customer Name</th>
                                        <th>Customer Mobile</th>
                                        <th>Customer Email</th>
                                    </tr>
                                </tfoot>
                                <tbody>
                                <?php 

                                    $query= $this->db->get('invoice');
                                    foreach ($query->result() as $row) {

                                    /*============================Use Postgresql php=========================*/
                                        // $tampil = pg_query("SELECT * FROM invoice") or die(pg_last_error());
                                        // while($result=pg_fetch_array($tampil)){
                                    /*=======================================================================*/

                                    /*============================USE PDO Postgres===========================*/
                                        // $sql = $myPDO->query("SELECT * FROM invoice") or die(pg_error());
                                        // while($result = $sql->fetch(PDO::FETCH_ASSOC)) {
                                    /*=======================================================================*/

                                ?>
                                    <tr class="odd gradeX">
                                        <td><?php echo $row->number; ?></td>
                                        <td><?php echo $row->invoice_date; ?></td>
                                        <td><?php echo $row->invoice_duedate; ?></td>
                                        <td><?php echo $row->currency_code; ?></td>
                                        <td><?php echo $row->total; ?></td>
                                        <td><?php echo $row->booking_number; ?></td>
                                        <td><?php echo $row->product_name; ?></td>
                                        <td><?php echo $row->status; ?></td>
                                        <td><?php echo $row->customer_name; ?></td>
                                        <td><?php echo $row->customer_mobile; ?></td>
                                        <td><?php echo $row->customer_email; ?></td>
                                        
                                    </tr>
                                 <?php } ?>   
                               </tbody>
                            </table>
                            <!-- /.table-responsive -->
                        </div>
                        <!-- /.panel-body -->
                    </div>
                    <!-- /.panel -->
                </div>
                <!-- /.col-lg-12 -->
            </div>     
            <!-- /.row -->
        </div>
        <!-- /#page-wrapper -->
        <a id="back-to-top" href="#" class="btn btn-primary btn-lg back-to-top" role="button" title="Click to return on the top page" data-toggle="tooltip" data-placement="left"><span class="glyphicon glyphicon-chevron-up"></span></a>
    </div>
    <!-- /#wrapper -->