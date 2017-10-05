 <!-- jQuery -->

    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="<?=base_url()?>assets/vendor/bootstrap/js/bootstrap.min.js"></script>

    <!-- Metis Menu Plugin JavaScript -->
    <script src="<?=base_url()?>assets/vendor/metisMenu/metisMenu.min.js"></script>
    
    <script src="<?=base_url()?>assets/vendor/Chart/Chart.bundle.js"></script>
    <script src="<?=base_url()?>assets/vendor/Chart/Chart.bundle.min.js"></script>
    <script src="<?=base_url()?>assets/vendor/Chart/Chart.js"></script>
    <script src="<?=base_url()?>assets/vendor/Chart/Chart.min.js"></script>
    <script src="<?=base_url()?>assets/data/chart-data-hg.js"></script>
    <script src="<?=base_url()?>assets/data/chart-data-eg.js"></script>
    <script type="text/javascript">
    function submitForm(action) {
        var form = document.getElementById('form1');
        if(form!=null){
            form.action = action;
            form.submit();
        }
        else{
        }
    }
    </script>

    <!-- Custom Theme JavaScript -->
    <script src="../dist/js/sb-admin-2.js"></script>