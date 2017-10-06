<?php
 
if (!defined('BASEPATH')) exit('No direct script access allowed');
 
/**
* Description of highchart_model
*
* @author http://www.roytuts.com
*/
class Chart_models extends CI_Model {
 
private $invoice = 'invoice';
 
/**
* get chart data
*/
function get_chart_invoice($start_date, $end_date) {

    
    $sql = "SELECT SUM(total) total, DATE(invoice_date) invoice_date
    FROM " . $this->invoice . "
    WHERE product_name='Home Guards' AND DATE(invoice_date) >= " . $this->db->escape($start_date) . "
    AND DATE(invoice_date) <= " . $this->db->escape($end_date) . "
    GROUP BY DATE(invoice_date) ORDER BY DATE(invoice_date) DESC";
    $query = $this->db->query($sql);
        if ($query->num_rows() > 0) {
            $data = array();
            foreach ($query->result_array() as $key => $value) {
                $data[$key]['label'] = $value['invoice_date'];
                $data[$key]['value'] = $value['total'];
            }
            return $data;
        }
        return NULL;
    } 
}
 
/* End of file highchart_model.php */
/* Location: ./application/models/highchart_model.php */