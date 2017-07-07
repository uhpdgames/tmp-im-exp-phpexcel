<?php
/***
██╗   ██╗██╗  ██╗██████╗ ██████╗  ██████╗  █████╗  ███╗   ███╗███████╗███████╗
██║   ██║██║  ██║██╔══██╗██╔══██╗██╔════╝ ██╔══██╗████╗ ████║██╔════╝██╔════╝
██║   ██║███████║██████╔╝██║  ██║██║  ███╗███████║██╔████╔██║█████╗  ███████╗
██║   ██║██╔══██║██╔═══╝ ██║  ██║██║   ██║██╔══██║██║ ╚██╔╝██║██╔══╝  ╚════██║
╚██████╔╝██║  ██║██║     ██████╔╝╚██████╔╝██║  ██║██║  ╚═╝ ██║███████╗███████║
╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═════╝   ╚═════╝ ╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝╚══════╝
@author: Ung Hoang Phi Dang
@version    0.01-beta, 2017/05
@copyright  http://uhpdgames.me/ - LGPL
@page-home  https://uhpdgames.github.io
 https://github.com/uhpdgames/tmp-im-exp-phpexcel
 *
 * @package    EXPORT/ IMPORT
 * @category   UHPD Games Dev
 */
/**
 * Request  PHPExcel
 * https://github.com/PHPOffice/PHPExcel
 */
//defined('PHP_EXCEL_LIBRARIES') or define('PHP_EXCEL_LIBRARIES', libraries_get_path('PHPExcel') . 'Classes/PHPExcel.php');
defined('PHP_EXCEL_LIBRARIES') or define('PHP_EXCEL_LIBRARIES', DRUPAL_ROOT . '/sites/all/libraries/PHPExcel/Classes/PHPExcel.php');

require_once PHP_EXCEL_LIBRARIES;
if(!class_exists('PHPExcel')){
  die('Not found! path ...Libs/PHPExcel/Classes/PHPExcel.php');
}

class UHPDGAMES_EXPORT{
  private $data;
  public $path;
  public $filename;
  public $auto_sweet;
  public $export_readonly = False;
  public $setLoadSheet = array('Sheet1');
  //
  private $im_values;
  /*  public function __construct(
      $data,
      $path = null,
      $filename = null
    ){
      $this->data = $data;
      $this->path = $path;
      $this->filename = $filename;
    }*/
  public function SET_data($data) {
    if(is_array($data)) return $this->data = $data;
    else return $this->data = array();
  }
  private function SET_template_values($data) {
    if(is_array($data)) return $this->im_values = serialize($data);
    else return $this->im_values = '';
  }
  public function GET_template_values() {
    if(!empty($this->im_values)) return unserialize($this->im_values);
  }

  /**
   * How to use
   * $this->SET_data(data)
   * export_excel()
   * @return bool
   */
  function export_excel(){
    try{
      $objPHPExcel = new PHPExcel();

      $data = isset($this->data) ? $this->data : array();

      if (count($data) <= 0) die(t('Không có dữ liệu để xuất.'));
      if(isset($this->path)) $path = $this->path;
      else $path = variable_get('file_public_path', conf_path() . '/files') . '/exports';

      /*if (!is_dir($path)) mkdir($path);
      if (!is_readable($path)) die('Không thể xuất dữ liệu, vui lòng kiểm tra lại phân quyền.');*/
      $filename = $this->filename;
      if (!isset($filename)) $filename = date('Y-m-d') . '.xlsx';
      else{
        $tmp_name = explode('.', $this->filename);
        if(count($tmp_name) > 1){
          if($tmp_name[1] =! 'xls' || $tmp_name[1] =! 'xls'){
            //die('Lỗi định dạng tệp xuất. E/x: name.xls');
            $filename = $tmp_name[0] .date('Y-m-d') .' '.'.xlsx';
          }
        }else{
          $filename .= date('Y-m-d') .' ' . '.xlsx';
        }
      }
      $path .= DIRECTORY_SEPARATOR . $filename;

      if($this->export_readonly) $objPHPExcel->setReadDataOnly(true);

      $active_sheet = $objPHPExcel->getActiveSheet();

      //title
      $sheet = array('columns' => 0, 'rows' => 2);
      if (isset($data['header'])) $this->set_values_export_excel($data['header'] , $active_sheet, $sheet, 'header');

      //values
      if(isset($data)) $this->set_values_export_excel($data, $active_sheet, $sheet, 'data');

      //custom
      if(isset($data['custom'])) {
        $custom = $data['custom'];

        $sheet = array('columns' => 0, 'rows' => $sheet['rows']);
        if(isset($custom['title'])) $this->set_values_export_excel($custom['title'], $active_sheet, $sheet, 'header');
        if(isset($custom['values'])) $this->set_values_export_excel($custom['values'], $active_sheet, $sheet, 'data');
      }
      /*more values...
        ...
      */

      if(isset($this->auto_sweet)) $this->auto_sweet_export_excel($active_sheet);
      //save
      $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
      ob_end_clean();
      $objWriter->save($path);
      if($path){
        drupal_add_http_header('Pragma', 'public');
        drupal_add_http_header('Expires', '0');
        drupal_add_http_header('Cache-Control', 'must-revalidate, post-check=0, pre-check=0');
        drupal_add_http_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        drupal_add_http_header('Content-Disposition', 'attachment; filename=' . basename($path) . ';');
        drupal_add_http_header('Content-Transfer-Encoding', 'binary');
        drupal_add_http_header('Content-Length', filesize($path));
        readfile($path);
        unlink($path);
        drupal_exit();
      }
      return true;
    }catch (Exception $exc){
      //var_dump($exc);
      die('Có lỗi xảy ra, có thể bộ nhớ không đủ để xử lý vui lòng kiểm tra lại.');
    }
    return false;
  }

  /**
   * How to use
   * [VALUES] [ROW][COLUMN] : A = 0 , B = 1, ....
   * import_template();
   * $this->GET_template_values();
   * @return bool
   */
  public function import_template() {
    if(isset($this->path)) $path = $this->path;
    else $path = variable_get('file_public_path', conf_path() . '/files') . '/uploads';

    if(isset($this->filename)) $filename = $this->filename;
    else die('Error! uploaded not working...');

    if(is_dir($filename)){
      $path = $filename;
    }else{
      $path .= DIRECTORY_SEPARATOR . $filename;
    }

    //debug
    //$path = 'C:\Bitnami\drupal-7.54-0\apps\drupal\htdocs\sites\all\modules\custom\diemrl\DEBUG\import\Import_nhapds_mau.xlsx';
    //$path = 'C:\Bitnami\drupal-7.54-0\apps\drupal\htdocs\sites\all\modules\custom\diemrl\DEBUG\Import_nhapds_mau.xlsx';
    try {
      $result = array();

      $objReader = new PHPExcel_Reader_Excel2007();
      $objReader->setReadDataOnly(true);
      $objReader->setLoadSheetsOnly( $this->setLoadSheet );

      $objPHPExcel = $objReader->load($path);
      //$objSheet = $objPHPExcel->getActiveSheet();
      //$result['getHighestRow'] = $objSheet->getHighestRow();

      foreach ($objPHPExcel->setActiveSheetIndex(0)->getRowIterator() as $row => $rows) {
        $cellIterator = $rows->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        foreach ($cellIterator as $cell) {
          if (!is_null($cell)) {
            $result['values'][$row][] = $cell->getCalculatedValue();
          }
        }
      }
      $this->SET_template_values($result);
      ob_end_clean();
      unlink($path);
      return True;
    }catch (Exception $exc){
      return False;
      die($exc);
    }
  }

  /**
   * Bố cục của excel
   * @param $active_sheet
   */
  private function auto_sweet_export_excel(&$active_sheet) {
    $col = array();
    foreach (range('A','Z') as $item) array_push($col, $item);
    foreach ($col as $item){
      $active_sheet ->getHeaderFooter()->setOddFooter('&L&B'. $item);
      $active_sheet ->getColumnDimension($item)->setAutoSize(true);
    }
    $active_sheet->getStyle('A1:Z1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
      ->getStartColor()->setARGB('FFFF0000');
  }

  /**
   * Ghi nhận kết qảy
   * @param $active_sheet
   * @param $sheet
   * @param $values
   * @param null $options
   */
  private function set_values_export_excel($values, &$active_sheet, &$sheet, $options = null) {
    $columns = 0;
    $rows = 2;

    if(isset($sheet['columns'])) $columns = $sheet['columns'];
    if(isset($sheet['rows'])) $rows = $sheet['rows'];

    if(count($values) > 0) {
      foreach ($values as $value) {
        switch ($options) {
          case 'header':
            $active_sheet->setCellValueExplicitByColumnAndRow($columns, 1, $value['data'], PHPExcel_Cell_DataType::TYPE_STRING);
            $columns++;
            break;
          case 'data':
            foreach ($value as $data) $active_sheet->setCellValueExplicitByColumnAndRow($columns++, $rows, $data, PHPExcel_Cell_DataType::TYPE_STRING);
            $columns = 0;
            $rows++;
            break;
          default;return;
        }
      }
    }
  }
}
?>