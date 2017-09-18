
<?php
	$this->load->helper('url');

	base_url('assets/vendor/vendor/autoload.php');

	use Raulr\GooglePlayScraper\Scraper;

	$scraper = new Scraper();
	$app = $scraper->getApp('org.apache.cordova.Mytra');
	echo $app['downloads'];
	
?>  