<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Highchart example in Codeigniter </title>
<link rel="stylesheet" href="<?php echo base_url(); ?>assets/css/charts/chart.css">
<link rel="stylesheet" href="<?php echo base_url(); ?>assets/css/charts/xcharts.min.css">

<link rel="stylesheet" href="<?php echo base_url(); ?>assets/css/daterangepicker.css">
<script src="<?php echo base_url(); ?>assets/js/jquery-1.9.1.min.js"></script>
<script src="<?php echo base_url(); ?>assets/js/jquery-migrate-1.2.1.js"></script>
<script src="<?php echo base_url(); ?>assets/js/jquery-ui-1.10.3-custom.min.js"></script>
<script type='text/javascript' src='<?php echo base_url(); ?>assets/js/charts/d3.v2.js'></script>
<script type='text/javascript' src='<?php echo base_url(); ?>assets/js/charts/sugar.min.js'></script>
<script type='text/javascript' src='<?php echo base_url(); ?>assets/js/charts/xcharts.min.js'></script>
<script type='text/javascript' src='<?php echo base_url(); ?>assets/js/charts/script.js'></script>
<script type='text/javascript' src='<?php echo base_url(); ?>assets/js/daterangepicker.js'></script>

    
    <link href="<?=base_url('assets/vendor/bootstrap/css/bootstrap.min.css')?>" rel="stylesheet">

    <!-- MetisMenu CSS -->
    <link href="<?=base_url('assets/vendor/metisMenu/metisMenu.min.css')?>" rel="stylesheet">

    <!-- Custom CSS -->
    <link href="<?=base_url('assets/dist/css/sb-admin-2.css')?>" rel="stylesheet">
    <link rel="stylesheet"  href="<?=base_url('assets/dist/css/scroll.css')?>">
    <!-- Custom Fonts -->
    <link href="<?=base_url('assets/vendor/font-awesome/css/font-awesome.min.css')?>" rel="stylesheet" type="text/css">

    <!-- Bootstrap Core JavaScript -->
    <script src="<?=base_url()?>assets/vendor/bootstrap/js/bootstrap.min.js"></script>

    <!-- Metis Menu Plugin JavaScript -->
    <script src="<?=base_url()?>assets/vendor/metisMenu/metisMenu.min.js"></script>
</head>
<body>
 <div id="wrapper">

        <!-- Navigation -->
        <nav class="navbar navbar-default navbar-static-top" role="navigation" style="margin-bottom: 0">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
                    <span class="sr-only">Toggle navigation</span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                    <span class="icon-bar"></span>
                </button>
                <a class="navbar-brand" href="<?=base_url('dashboard')?>"><img src="<?=base_url() ?>assets/vendor/images/mytra-logo.png" height="30" width="30" alt="logo"></a>
                <a class="navbar-brand" href="<?=base_url('dashboard')?>">mytra Dashboard</a>
            </div>
            <!-- /.navbar-header -->

            <ul class="nav navbar-top-links navbar-right">
                <li class="dropdown">
                    <a class="dropdown-toggle" data-toggle="dropdown" href="#">
                        <i class="fa fa-envelope fa-fw"></i> <i class="fa fa-caret-down"></i>
                    </a>
                    <ul class="dropdown-menu dropdown-messages">
                        <li>
                            <a href="#">
                                <div>
                                    <strong>John Smith</strong>
                                    <span class="pull-right text-muted">
                                        <em>Yesterday</em>
                                    </span>
                                </div>
                                <div>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque eleifend...</div>
                            </a>
                        </li>
                        <li class="divider"></li>
                        <li>
                            <a class="text-center" href="#">
                                <strong>Read All Messages</strong>
                                <i class="fa fa-angle-right"></i>
                            </a>
                        </li>
                    </ul>
                    <!-- /.dropdown-messages -->
                </li>
                <!-- /.dropdown -->               
                <li class="dropdown">
                    <a class="dropdown-toggle" data-toggle="dropdown" href="#">
                        <i class="fa fa-bell fa-fw"></i> <i class="fa fa-caret-down"></i>
                    </a>
                    <ul class="dropdown-menu dropdown-alerts">
                        <li>
                            <a href="#">
                                <div>
                                    <i class="fa fa-comment fa-fw"></i> New Comment
                                    <span class="pull-right text-muted small">4 minutes ago</span>
                                </div>
                            </a>
                        </li>                          
                    </ul>
                    <!-- /.dropdown-alerts -->
                </li>
                <!-- /.dropdown -->
                <li class="dropdown">
                    <a class="dropdown-toggle" data-toggle="dropdown" href="#">
                        <i class="fa fa-user fa-fw"></i> <i class="fa fa-caret-down"></i>
                    </a>
                    <ul class="dropdown-menu dropdown-user">
                        <li><a href="#"><i class="fa fa-user fa-fw"></i> User Profile</a>
                        </li>
                        <li><a href="#"><i class="fa fa-gear fa-fw"></i> Settings</a>
                        </li>
                        <li class="divider"></li>
                        <li><a href="<?=base_url('login/logout');?>"><i class="fa fa-sign-out fa-fw"></i> Logout</a>
                        </li>
                    </ul>
                    <!-- /.dropdown-user -->
                </li>
                <!-- /.dropdown -->
            </ul>
            <!-- /.navbar-top-links -->

            <div class="navbar-default sidebar" role="navigation">
                <div class="sidebar-nav navbar-collapse">
                    <ul class="nav" id="side-menu">
                        <li class="sidebar-search">
                            <div class="input-group custom-search-form">
                                <input type="text" class="form-control" placeholder="Search...">
                                <span class="input-group-btn">
                                <button class="btn btn-default" type="button">
                                    <i class="fa fa-search"></i>
                                </button>
                            </span>
                            </div>
                            <!-- /input-group -->
                        </li>
                        <li>
                            <a href="<?=base_url('dashboard');?>"><i class="fa fa-dashboard fa-fw"></i> Dashboard</a>
                        </li>
                        <li>
                            <a href="#"><i class="fa fa-bar-chart-o fa-fw"></i> Charts<span class="fa arrow"></span></a>
                            <ul class="nav nav-second-level">
                                <li>
                                    <a href="home-guards">Home Guards</a>
                                </li>
                                <li>
                                    <a href="<?=base_url('dashboard/event_chart');?>">Event Guards</a>
                                </li>
                            </ul>
                            <!-- /.nav-second-level -->
                        </li>
                        <li>
                            <a href="#"><i class="fa fa-table fa-fw"></i> Tables<span class="fa arrow"></span></a>
                            <ul class="nav nav-second-level">
                                <li>
                                    <a href="<?php echo base_url('tables/booking') ;?>">Booking</a>
                                </li>
                                <li>
                                    <a href="<?php echo base_url('tables/invoice'); ?>">Invoice</a>
                                </li>
                                <li>
                                    <a href="<?php echo base_url('tables/customer'); ?>">Customer</a>
                                </li>
                            </ul>
                        </li>                       
                    </ul>
                </div>
                <!-- /.sidebar-collapse -->
            </div>
            <!-- /.navbar-static-side -->
        </nav>
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
                        <form class="form-horizontal">
                            <fieldset>
                                <div class="input-prepend">
                                <input type="text" name="range" id="range" class="pull-left" style="padding: 5px 10px; border: 1px solid #ccc; width: 90%" ><i class="fa fa-calendar"></i>&nbsp;<span class="add-on"></span>
                                    
                                </div>
                            </fieldset>
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
                            <div id="msg"></div>
                            <figure id="chart"></figure>
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
</body>
</html>