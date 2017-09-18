<?php

	function setInterval($f, $milliseconds)
{
    $seconds=(int)$milliseconds/1000;
    while(true)
    {
        $f();
        sleep($seconds);
    }
}
	

	setInterval(function(){

	date_default_timezone_set('Australia/Melbourne');
	$date = date('m/d/Y h:i:s a', time());

	$date = strtotime($date);
	$lala = '1';
	while ($lala<11){
    echo date('%d ', $lala);
    $lala = $lala++;
}
}, 1000);
			
?>