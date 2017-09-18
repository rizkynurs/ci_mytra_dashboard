
        <div id="page-wrapper">
            <div class="row">
                <div class="col-lg-12">
                    <h1 class="page-header">Customer's Tables
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
                            Customer's Tables
                            <div class="pull-right">
                            <div class="btn-group">
                                    <button type="button" class="btn btn-default btn-xs dropdown-toggle" data-toggle="dropdown">
                                        Actions
                                        <span class="caret"></span>
                                    </button>
                                    <ul class="dropdown-menu pull-right" role="menu">
                                        <li><a href="export-customer">Download to EXCEL</a>
                                        </li>
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
                                        <th>Name</th>
                                        <th>Email</th>
                                        <th>Phone</th>
                                        <th>Mobile Phone</th>
                                    </tr>
                                </thead>
                                <tfoot>
                                    <tr>
                                        <th>Name</th>
                                        <th>Email</th>
                                        <th>Phone</th>
                                        <th>Mobile Phone</th>
                                    </tr>
                                </tfoot>
                                <tbody>
                                <?php 
                                    /*============================Use Postgresql php=========================*/
                                        $tampil = pg_query("SELECT * FROM customer") or die(pg_last_error());
                                        while($result=pg_fetch_array($tampil)){
                                    /*=======================================================================*/

                                    /*============================USE PDO Postgres===========================*/
                                        // $sql = $myPDO->query("SELECT * FROM invoice") or die(pg_error());
                                        // while($result = $sql->fetch(PDO::FETCH_ASSOC)) {
                                    /*=======================================================================*/

                                ?>
                                    <tr> 
                                        <td><?php echo $result['name']; ?></td>
                                        <td><?php echo $result['email']; ?></td>
                                        <td><?php echo $result['phone']; ?></td>
                                        <td><?php echo $result['mobile']; ?></td>
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

  