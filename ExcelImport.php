<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2017/9/20
 * Time: 16:09
 * Author she
 * Content Excel import
 */
namespace app\models\events;

use app\models\CustomerChildrenGrade;
use app\models\CustomerInfo;
use app\models\CustomerSource;
use app\models\IntegralDegree;
use app\models\Upload;
use moonland\phpexcel\Excel;
use PHPExcel_Reader_Excel5;
use yii\web\UploadedFile;
use app\models\froms\StudentInfoImportForm;
use OSS\OssClient;
use Yii;
use yii\helpers\ArrayHelper;
use yii\data\ActiveDataProvider;
use app\core\base\BaseActiveRecord;
use PHPExcel;
use PHPExcel_Style_NumberFormat;

class ExcelImport extends \app\core\base\BaseModel
{
    /**
     * @param $string
     * @return int
     * 验证手机号
     */
    public static function CheckPhone($string){
        //检查号码格式是否错误
        if(!ExcelImport::Regular($string)){
            //echo('手机号码格式错误');
            //throw new \Exception('手机号码格式错误！！！');
            $data = ['word'=>'-1','data'=>'手机号码格式错误'];
            return $data;
        }
        //检查手机号是否已经在customer_info中
        if(!ExcelImport::Exists($string)){
            //echo('手机号码已存在');
            $data = ['word'=>'-2','data'=>'手机号码已存在'];
            return $data;
        }
        return 2;
    }
    /**
     * @param $string
     * @return bool
     * 手机号的正则匹配
     */
    public static function Regular($string){
        $rule  = "/^[1][345678][0-9]{9}$/";
        if(preg_match($rule,$string)){
            return true;
        }else{
            return false;
        }
    }

    /**
     * @param $string
     * @return bool
     * 判断手机号码是否在customer_info中存在
     */
    public static function Exists($string){
        $num = CustomerInfo::find()->where(['phone_number'=>$string])->count();
        if($num<1){
            unset($num);
            return true;
        }else{
            return false;
        }
    }

    /**
     * @param $model
     * @return string
     * 保存Excel文件
     */
    public static function SaveFile($model,$type){
        header("Content-Type:text/html;charset=utf-8");
//        $file = UploadedFile::getInstances($model,'education_certification')[0];  //获取上传的文件实例
//        if ($file) {
//            $filename = date('Y-m-d', time()) . '_' . rand(1, 9999999) . "." . $file->extension;//文件根据上传日期命名
//            $folder = self::setFolderExcel();
//            $file->saveAs($folder.$filename);   //保存文件
//            $format = $file->extension;
//        }
//        if(in_array($format,array('xls','xlsx'))) {
//            $excelFile = $folder.$filename . '';//获取文件名
//        }
        $upload = new Upload();
        $fileName = $upload->uploadExcel($model, 'education_certification',$type);
        $tag_data = self::NewExcel($fileName);
//        $tag_data = Excel::import($fileName, [
//            'setFirstRecordAsKeys' => true,
//            'setIndexSheetByName' => true,
//            //'getOnlySheet' => 'sheet1',
//        ]);
        return $tag_data;
    }

    /**
     * @param string $path
     * @return string
     */
    public function setFolderExcel($path="upload"){
        //$root = dirname(dirname(__FILE__))."/";
        $folder = str_replace("\\","/",dirname(dirname(dirname(__FILE__))))."/".$path."/".date('Ymd')."/";
        //$folder = "./".$path."/".date('Ymd')."/";
        if(!is_dir($folder))
        {
            if(!mkdir($folder, 0777, true)){
                die("创建目录失败");
            }else{
                chmod($folder,0777);
            }
        }
        return $folder;
    }

    /**
     * @param $arr
     * @return bool
     * 导入Excel资源到数据库
     */
    public static function Imports($arr){
    }

    /**
     * @param $tag_data
     * @throws \yii\db\Exception
     * @content 将所有数据每100条使用事务进行保存一次
     */
    public static function EachImport($tag_data){
        echo "************保存开始**************</br>";
        //$secondArray = array_chunk(array_slice($tag_data,1),100);
        foreach(array_slice($tag_data,1) as $key=>$data){
            $transaction = \Yii::$app->db->beginTransaction();
            try{
                //foreach($thirdArray as $key=>$data){
                    if(!empty($data['1'])){
                            $phone = $data['1'];
                            if(ExcelImport::CheckPhone($phone)!=2){
                                echo ExcelImport::CheckPhone($phone)['data']."</br>";
                            }elseif(!ExcelImport::Imports($data)){
                                //将对应的数据存到表中
                                echo "保存失败</br>";
                            }
                    }else{
                        continue;
                    }
                //}
                $transaction->commit();
                unset($data);
            }catch (\Exception $e){
                pr($e);
                $transaction->rollBack();
            }
        }
        echo "************保存结束**************</br>";
        echo "<hr>";
    }

    public static function NewExcel($filename,$encode='utf-8'){
        $objReader = \PHPExcel_IOFactory::createReader('Excel2007');
        $objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($filename);
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $highestRow = $objWorksheet->getHighestRow();
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = \PHPExcel_Cell::columnIndexFromString($highestColumn);
        $excelData = array();
        for ($row = 1; $row <= $highestRow; $row++) {
            for ($col = 0; $col < $highestColumnIndex; $col++) {
                $excelData[$row][] =(string)$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
            }
        }
        return $excelData;
    }

    /**
     * 公共导出模板
     */
    public static function Export()
    {
        $headerArr = ['姓名','手机号码','性别','年龄','访问设备','目前学历','孩子年级','备注'];

        $fileName = date('Y-m-d',time()).'_customer_'.rand(1,9999999);
        $objPHPExcel = new PHPExcel();
        $objProps = $objPHPExcel->getProperties();

        $key = ord('A');
        foreach($headerArr as $v){
            $colum = chr($key);
            $objPHPExcel->setActiveSheetIndex(0)->setCellValue($colum.'1',$v);
            $key += 1;
        }
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('R1','子女年级');
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue('Q1','学历');

        for($i=2;$i<100;$i++){
            $keyFirst = ord('A');
            foreach($headerArr as $v){
                $colum = chr($keyFirst);
                $objPHPExcel->getActiveSheet()->setCellValue($colum.$i,'');
                self::Explain($objPHPExcel,$i);
                $keyFirst += 1;
            }
        }

        $objPHPExcel->getActiveSheet()->setTitle('sheet1');
        $objPHPExcel->setActiveSheetIndex(0);
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(10);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(10);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(10);
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('R')->setWidth(15);
        $objPHPExcel->getActiveSheet()->getColumnDimension('Q')->setWidth(15);

        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('D1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('D1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('D1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('E1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('E1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('E1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('F1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('F1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('F1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('G1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('G1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('G1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('H1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('H1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('H1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('R1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('R1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('R1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('Q1')->getFont()->setName('黑体');
        $objPHPExcel->getActiveSheet()->getStyle('Q1')->getFont()->setSize(11);
        $objPHPExcel->getActiveSheet()->getStyle('Q1')->getFont()->setBold(true);

        $objPHPExcel->getActiveSheet()->getStyle('B')->getNumberFormat()
            ->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_TEXT);


        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header("Content-Disposition: attachment; filename=\"$fileName\".xls");
        header('Cache-Control: max-age=0');

        $writer = \PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel2007');
        $writer->save('php://output');
        /*if($ex == '2007') { //导出excel2007文档
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header("Content-Disposition: attachment; filename=\"$fileName\".xlsx");
            header('Cache-Control: max-age=0');
            $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
            $objWriter->save('php://output');
            exit;
        } else {  //导出excel2003文档
            header('Content-Type: application/vnd.ms-excel');
            header("Content-Disposition: attachment; filename=\"$fileName\".xls");
            header('Cache-Control: max-age=0');
            $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
            $objWriter->save('php://output');
            exit;
        }*/
    }

    /**
     * @param $objPHPExcel
     * 下拉选项
     */
    public static function Explain($objPHPExcel,$i){
        //性别
        $objValidationC = $objPHPExcel->getActiveSheet()->getCell(chr(ord('C')).$i)->getDataValidation(); //这一句为要设置数据有效性的单元格
        $objValidationC -> setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
            -> setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
            -> setAllowBlank(false)
            -> setShowInputMessage(true)
            -> setShowErrorMessage(true)
            -> setShowDropDown(true)
            -> setErrorTitle('输入的值有误')
            -> setError('您输入的值不在下拉框列表内.')
            -> setPromptTitle('性别')
            -> setFormula1('"男,女,未知"');

        //访问设备
        $objValidationE = $objPHPExcel->getActiveSheet()->getCell(chr(ord('E')).$i)->getDataValidation(); //这一句为要设置数据有效性的单元格
        $objValidationE -> setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
            -> setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
            -> setAllowBlank(false)
            -> setShowInputMessage(true)
            -> setShowErrorMessage(true)
            -> setShowDropDown(true)
            -> setErrorTitle('输入的值有误')
            -> setError('您输入的值不在下拉框列表内.')
            -> setPromptTitle('访问设备')
            -> setFormula1('"未知,移动端,PC端"');

        //来源
//        $source = ArrayHelper::getColumn(User::find()->asArray()->all(),'user_name');
//        $sourceE = implode(',',$source);
//        //$valueE = '"'.$sourceE.'"';
//        $str_len_e = strlen($sourceE);
//        if($str_len_e>=255){
//            $str_list_arr = explode(',', $sourceE);
//            if($str_list_arr)
//                foreach($str_list_arr as $i =>$d){
//                    $c = "P".($i+1);
//                    $objPHPExcel->getActiveSheet()->setCellValue($c,$d);
//                }
//            $endcellE = $c;
//            $objPHPExcel->getActiveSheet()->getColumnDimension('P')->setVisible(false);
//        }
//        $objValidationF = $objPHPExcel->getActiveSheet()->getCell(chr(ord('F')).$i)->getDataValidation(); //这一句为要设置数据有效性的单元格
//        $objValidationF -> setType(\PHPExcel_Cell_DataValidation::TYPE_LIST)
//            -> setErrorStyle(\PHPExcel_Cell_DataValidation::STYLE_INFORMATION)
//            -> setAllowBlank(false)
//            -> setShowInputMessage(true)
//            -> setShowErrorMessage(true)
//            -> setShowDropDown(true)
//            -> setErrorTitle('输入的值有误')
//            -> setError('您输入的值不在下拉框列表内.')
//            -> setPromptTitle('信息来源');
//            if($str_len_e<255){
//                $objValidationF->setFormula1('"'.$sourceE.'"');
//            }else{
//                $objValidationF->setFormula1("sheet1!P1:{$endcellE}");
//            }
    }

    public static function NewImport($arr){
    }
}
