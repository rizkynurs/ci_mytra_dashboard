 <div id="page-wrapper">
            <div class="row">
                <div class="col-lg-12">
                    <h1 class="page-header">Dashboard</h1>
                </div>
                <!-- /.col-lg-12 -->
            </div>
            <!-- /.row -->
            <div class="row">
                <div class="col-lg-3 col-md-6">
                    <div class="panel panel-primary">
                        <div class="panel-heading">
                            <div class="row">
                                <div class="col-xs-3">
                                    <i class="fa fa-calendar fa-4x"></i>
                                </div>
                                <div class="col-xs-9 text-right">
                                    <div class="huge">
                                        <?php $this->load->view('pages/dashboard/booking-number'); ?>
                                    </div>
                                    <div>Booking</div>
                                </div>
                            </div>
                        </div>
                        <a href="<?=base_url('tables/booking');?>">
                            <div class="panel-footer">
                                <span class="pull-left">View Details</span>
                                <span class="pull-right"><i class="fa fa-arrow-circle-right"></i></span>
                                <div class="clearfix"></div>
                            </div>
                        </a>
                    </div>
                </div>
                <div class="col-lg-3 col-md-6">
                    <div class="panel panel-green">
                        <div class="panel-heading">
                            <div class="row">
                                <div class="col-xs-3">
                                    <i class="fa fa-money fa-4x"></i>
                                </div>
                                <div class="col-xs-9 text-right">
                                    <div class="huge"> 
                                        <?php $this->load->view('pages/dashboard/invoice-number.php') ?> 
                                    </div>
                                    <div>Income (in Million)</div>
                                </div>
                            </div>
                        </div>
                        <a href="<?=base_url('tables/invoice')?>">
                            <div class="panel-footer">
                                <span class="pull-left">View Details</span>
                                <span class="pull-right"><i class="fa fa-arrow-circle-right"></i></span>
                                <div class="clearfix"></div>
                            </div>
                        </a>
                    </div>
                </div>
                <div class="col-lg-3 col-md-6">
                    <div class="panel panel-yellow">
                        <div class="panel-heading">
                            <div class="row">
                                <div class="col-xs-3">
                                    <i class="fa fa-users fa-4x"></i>
                                </div>
                                <div class="col-xs-9 text-right">
                                    <div class="huge">
                                        <?php $this->load->view('pages/dashboard/customer-number');?>
                                    </div>
                                    <div>User Sign Up</div>
                                </div>
                            </div>
                        </div>
                        <a href="#">
                            <div class="panel-footer">
                                <span class="pull-left">View Details</span>
                                <span class="pull-right"><i class="fa fa-arrow-circle-right"></i></span>
                                <div class="clearfix"></div>
                            </div>
                        </a>
                    </div>
                </div>
                <div class="col-lg-3 col-md-6">                           
                        <div class="panel panel-red">
                            <div class="panel-heading-1">
                                <div class="row">
                                    <div class="col-xs-3">
                                        <i class="fa fa-download fa-2x"></i>
                                    </div>
                                    <div class="col-xs-9 text-right">
                                        <div>
                                            <?php $this->load->view('pages/dashboard/customer-number'); ?>
                                        </div>
                                        <div>Downloads</div>
                                    </div>
                                </div>
                            </div>
                            <div class="panel-heading-1">
                                <div class="row">
                                    <div class="col-xs-3">
                                        <i class="fa fa-eye fa-2x"></i>
                                    </div>
                                    <div class="col-xs-9 text-right">
                                        <div>
                                            <?php $this->load->view('pages/dashboard/invoice-number'); ?>
                                        </div>
                                        <div>Views</div>
                                    </div>
                                </div>
                            </div>
                            <div class="panel-heading-1">
                                <div class="row">
                                    <div class="col-xs-3">
                                        <i class="fa fa-sign-in fa-2x"></i>
                                    </div>
                                    <div class="col-xs-9 text-right">
                                        <div>
                                            <?php $this->load->view('pages/dashboard/customer-number'); ?>
                                        </div>
                                        <div>Log-in</div>
                                    </div>
                                </div>
                            </div>  
                            <a href="#">
                                <div class="panel-footer">
                                    <span class="pull-left">View Details</span>
                                    <span class="pull-right"><i class="fa fa-arrow-circle-right"></i></span>
                                    <div class="clearfix"></div>
                                </div>
                            </a>
                        </div>                
                </div>
            </div>
            <!-- /.row -->
            <div class="row">
                <div class="col-lg-12">            
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            Pick a Date
                        </div>
                        <!-- /.panel-heading -->
                        <div class="panel-body">
                            <div class="white-section" style="padding:20px">
                                <span class="montserrat-text uppercase" style="font-size:24px">Pick a date</span>
                        <p>
                            Choose a date of the week that you want to see.
                        </p>
                        <form id="form1" method="post">
                            <div class="input_1">
                                <input type="text" id="pass_date" name="pass_date" >
                                <span class="text-muted" style="font-size:10px">YYYY-MM-DD</span>       
                            </div>
                            <input type="submit" onclick="submitForm('home-guards')" value="Home Guards" class="btn green" id="hgsubmit" style="margin-top:20px"></input>
                            <input type="submit" onclick="submitForm('event-guards')" value="Event Guards" class="btn green" id="egsubmit" style="margin-top:20px"></input>
                                
                                <?php 
                                    if(isset($_POST['submit'])){
                                        $tfDate = $_POST['pass_date'];   
                                    }
                                    if(isset($tfDate)){ 
                                        echo $tfDate;
                                    }
                                ?>
                                
                        </form>
                            </div>
                                
                        </div>
                        <!-- /.panel-body -->
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="col-lg-12">
                    <div class="panel panel-default">
                        <div class="panel-heading">
                            <i class="fa fa-bar-chart-o fa-fw"></i> Daily Book
                            <div class="pull-right">
                                <div class="btn-group">
                                    <button type="button" class="btn btn-default btn-xs dropdown-toggle" data-toggle="dropdown">
                                        Actions
                                        <span class="caret"></span>
                                    </button>
                                    <ul class="dropdown-menu pull-right" role="menu">
                                        <li><a href="#">Action</a>
                                        </li>
                                        <li><a href="#">Another action</a>
                                        </li>
                                        <li><a href="#">Something else here</a>
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
                            <canvas id="hgbookdaily" width="400" height="300"></canvas>                           
                            <canvas id="egbookdaily" width="400" height="300"></canvas>                                                         
                        </div>
                        <!-- /.panel-body -->
                    </div>
                    <!-- /.panel -->                    
                </div>
                <!-- /.col-lg-8 -->
            </div>
            <!-- /.row -->
        <a id="back-to-top" href="#" class="btn btn-primary btn-lg back-to-top" role="button" title="Click to return on the top page" data-toggle="tooltip" data-placement="left"><span class="glyphicon glyphicon-chevron-up"></span></a>
        </div>
        <!-- /#page-wrapper -->

    </div>
    <!-- /#wrapper -->

    <script src="<?php echo base_url(); ?>assets/vendor/Chart/Chart.bundle.js"></script>
    <script src="<?php echo base_url(); ?>assets/vendor/Chart/Chart.bundle.min.js"></script>
    <script src="<?php echo base_url(); ?>assets/vendor/Chart/Chart.js"></script>
    <script src="<?php echo base_url(); ?>assets/vendor/Chart/Chart.min.js"></script>
    <script src="<?php echo base_url(); ?>assets/data/chart-data-hg.js"></script>
    <script src="<?php echo base_url(); ?>assets/data/chart-data-eg.js"></script>