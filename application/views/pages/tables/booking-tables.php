        <div id="page-wrapper">
            <div class="row">
                <div class="col-lg-12">
                    <h1 class="page-header">Booking Tables
                    </h1>
                </div>
                <!-- /.col-lg-12 -->
            </div>
            <!-- /.row -->
            <div class="row">
                <div class="col-lg-12">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            Home Guard's Tables
                            <div class="pull-right">
                            <div class="btn-group">
                                    <button type="button" class="btn btn-default btn-xs dropdown-toggle" data-toggle="dropdown">
                                        Actions
                                        <span class="caret"></span>
                                    </button>
                                    <ul class="dropdown-menu pull-right" role="menu">
                                        <li><a href="<?=base_url('export/booking');?>">Download All Data Booking to EXCEL</a></li>
                                        <li><a href="<?=base_url('export/booking_home_guards');?>">Download Data Home Guards to EXCEL</a></li>
                                            <li><a href="<?=base_url('export/booking_event_guards');?>">Download Event Guards to EXCEL</a>
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
                            <table width="100%" class="table table-striped table-bordered table-hover" id="dataTables-example" cellspacing="0">
                            <!-- <table width="100%" class="table table-striped table-bordered table-hover" id="dataTables-example"> -->
                                 <thead>
                                    <tr>
                                        <th class="center" align="center">ID</th>
                                        <th class="center" align="center">Address</th>
                                        <th class="center" align="center">PIC</th>
                                        <th class="center" align="center">Booking From</th>
                                        <th class="center" align="center">Booking Until</th>
                                        <th class="center" align="center">Quantity</th>
                                        <th class="center" align="center">Price</th>
                                        <th class="center" align="center">Location</th>
                                        <th class="center" align="center">Status</th>
                                    </tr>
                                </thead>
                                <tfoot>
                                    <tr>
                                        <th>ID</th>
                                        <th>Address</th>
                                        <th>PIC</th>
                                        <th>Booking From</th>
                                        <th>Booking Until</th>
                                        <th>Quantity</th>
                                        <th>Price</th>
                                        <th>Location</th>
                                        <th>Status</th>
                                    </tr>
                                </tfoot>
                                
                                <tbody>
                                <?php 
                                    $query = $this->db->get('booking');
                                    // $query = $this->db->get_where('booking',array('product_name'=>'Home Guards'));
                                    foreach ($query->result() as $row) {
                                    /*============================Use Postgresql php=========================*/
                                        // $tampil1 = pg_query("SELECT * FROM booking WHERE product_name='Home Guards'") or die(pg_error());
                                        // while($result=pg_fetch_array($tampil1)){
                                    /*=======================================================================*/

                                    /*============================USE PDO Postgres===========================*/
                                        // $sql = $myPDO->query("SELECT * FROM invoice") or die(pg_error());
                                        // while($result = $sql->fetch(PDO::FETCH_ASSOC)) {
                                    /*=======================================================================*/

                                ?>
                                    <tr class="odd grad eX">
                                        <td><?php echo $row->booking_id; ?></td>
                                        <td><?php echo $row->address; ?></td>
                                        <td><?php echo $row->pic; ?></td>
                                        <td><?php echo $row->booking_from; ?></td>
                                        <td><?php echo $row->booking_until; ?></td>
                                        <td class="center" align="center"><?php echo $row->qty; ?></td>
                                        <td><?php echo $row->price; ?></td>
                                        <td><?php echo $row->geolocation; ?></td>
                                        <td><?php echo $row->status; ?></td>
                                        
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
