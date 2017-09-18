<!doctype html>
<html>
<head>
 <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">
<link href="<?php echo base_url();?>assets/vendor/images/mytra-logo.png" rel="shortcut icon">
<title><?php echo $title ?></title>
 <!-- Bootstrap Core CSS -->
    <link href="<?php echo base_url() ?>assets/vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">

    <!-- MetisMenu CSS -->
    <link href="<?php echo base_url() ?>assets/vendor/metisMenu/metisMenu.min.css" rel="stylesheet">

    <!-- Custom CSS -->
    <link href="<?php echo base_url() ?>assets/dist/css/sb-admin-2.css" rel="stylesheet">

    <!-- Custom Fonts -->
    <link href="<?php echo base_url() ?>assets/vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
</head>

<body>
 <div class="container">
    <div class="row">
      <div class="col-md-4 col-md-offset-4">
        <div class="login-panel panel panel-default">
          <div class="panel-heading">
            <h2 class="panel-title" align="center"><img src="<?php echo base_url();?>assets/vendor/images/mytra-logo.png" rel="shortcut icon"></h2>
	          <h1 align="center">Login Page</h2>
          </div>
        <div class="panel-body">
    
         <?php
    	 // Cetak session
    	if($this->session->flashdata('sukses')) {
    		echo '<p class="warning" style="margin: 10px 20px;">'.$this->session->flashdata('sukses').'</p>';
    	}
    	// Cetak error
    	echo validation_errors('<p class="warning" style="margin: 10px 20px;">','</p>');
    	?>
        
        <form role="form" action="<?php echo base_url('login') ?>" method="post">
          <fieldset>
              <div class="form-group">
                  <input class="form-control" id="username" placeholder="Username" name="username" type="text" autofocus>
              </div>
              <div class="form-group">
                  <input class="form-control" id="password" placeholder="Password" name="password" type="password" value="">
              </div>
              <div class="checkbox">
                  <label>
                      <input name="remember" type="checkbox" value="Remember Me">Remember Me
                  </label>
                      </div>
                                    <!-- Change this to a button or input when using this as a form -->
                      <input type = "submit" id="submit" value = " Login " /><br />                                
          </fieldset>
        </form>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="<?php echo base_url() ?>assets/vendor/jquery/jquery.min.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="<?php echo base_url() ?>assets/vendor/bootstrap/js/bootstrap.min.js"></script>

    <!-- Metis Menu Plugin JavaScript -->
    <script src="<?php echo base_url() ?>assets/vendor/metisMenu/metisMenu.min.js"></script>

    <!-- Custom Theme JavaScript -->
    <script src="<?php echo base_url() ?>assets/dist/js/sb-admin-2.js"></script>
</body>
</html>
