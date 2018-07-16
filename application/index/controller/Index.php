<?php
namespace app\index\controller;
use think\Controller;
use think\loader;
use think\Db;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Cell;

class Index extends Controller
{
    public function index()
    {
    
     // $wp=Db::name("jiaoshi")->where('id',60)->find();
     // print_r($wp);die;
       $this->assign('domain',$this->request->url(true));
        return $this->fetch('index');
      
    }
    public function import(){
      return $this->fetch('index/import');
    }
    public function export(){
      $this->wukebiao();
      return $this->fetch('index/export');
    }

    //导入课程信息excel表
    public function inputKecheng()
    {
      $file = request()->file('file');

      if(empty($file)) {//如果文件存在
        return $this->error('请按照流程操作','Index');
        
      }
      
          
       // 引入PHPExcel
       header("content-type:text/html; charset=utf-8");
       vendor("PHPExcel.PHPExcel");

       
// 移动到框架应用根目录/public/uploads/ 目录下
$filename="";
if($file){
    $info = $file->move(ROOT_PATH . 'public' . DS . 'uploads');
    if($info){
        $filename=ROOT_PATH . 'public' . DS . 'uploads'.DS.$info->getSaveName();
    }else{
        // 上传失败获取错误信息
        return array("resultcode" => -4, "resultmsg" => "文件上传失败", "data" => $file->getError());
    }
}

if(file_exists($filename)) {//如果文件存在
	echo "<script >alert('文件上传成功');</script>";
}
// Db::name('jiaoshi')->delete(true);
Db::execute('TRUNCATE table wp_jiaoshi');
// print_r($filename);die;
       // 文件导入
       // $filename="E:\\wamp64\www\\thinkphp5\\public\\style\\uploads\\"."$filename".".xls";
       $filename=iconv('utf-8', 'gbk', $filename);
       if(file_exists($filename)){
       	// 如果文件存在
       	$PHPReader = new \PHPExcel_Reader_Excel5();
       	// 载入excel文件
       	$PHPExcel = $PHPReader->load($filename);
       	// print_r($PHPExcel);die;
       	$sheet = $PHPExcel->getActiveSheet(0);//获取sheet
       	
        $highestRow = $sheet ->getHighestRow();//获取共有数据数
         // print_r($highestRow);die;
      $data=$sheet->toArray();
        for($i=1;$i<$highestRow;$i++){
        	$user['xiaoqu']=$data[$i][0];
        	$user['lou']=$data[$i][1];
        	
        	$user['number']= $this->findNum($data[$i][2]);
        	$var=$this->jieciNum($data[$i][3]);
        	$user['xingqi'] =$var[0];
        	$user['jieci'] = $var[1];
        	// $modelUserLogic = Loader::model('User','logic');//可能是链接数据库
        	// $ret=$modelUserLogic->add($user);
        	// print_r($user);die;
        	$success=Db::name('jiaoshi')->insert($user); //批量插入数据
        }
        return $this->success('导入成功','index/export');
       }else{
       	     return  array("resultcode" => -5,"resultmsg" =>"文件不存在","data" =>null);
       }

    }
    //导入教室号
    public function inputNumber(){

      $file = request()->file('file');

          if(empty($file)) {//如果文件存在
            return $this->error('请按照流程操作','Index');

          }


       // 引入PHPExcel
          header("content-type:text/html; charset=utf-8");
          vendor("PHPExcel.PHPExcel");


// 移动到框架应用根目录/public/uploads/ 目录下
          $filename="";
          if($file){
            $info = $file->move(ROOT_PATH . 'public' . DS . 'uploads');
            if($info){
              $filename=ROOT_PATH . 'public' . DS . 'uploads'.DS.$info->getSaveName();
            }else{
        // 上传失败获取错误信息
              return array("resultcode" => -4, "resultmsg" => "文件上传失败", "data" => $file->getError());
            }
          }

          if(file_exists($filename)) {//如果文件存在
            echo "文件上传成功";
          }
          
          // print_r($filename);die;
          // 文件导入
          // $filename="E:\\wamp64\www\\thinkphp5\\public\\style\\uploads\\"."$filename".".xls";
         Db::execute('TRUNCATE table wp_xuhao');
          $filename=iconv('utf-8', 'gbk', $filename);
          if(file_exists($filename)){
            // 如果文件存在
            $PHPReader = new \PHPExcel_Reader_Excel5();
            // 载入excel文件
            $PHPExcel = $PHPReader->load($filename);
            // print_r($PHPExcel);die;
            $sheet = $PHPExcel->getActiveSheet(0);//获取sheet
        
            $highestRow = $sheet ->getHighestRow();//获取共有数据数
         // print_r($highestRow);die;
            $data=$sheet->toArray();
            for($i=1;$i<$highestRow;$i++){
              $user['xuhao']=$data[$i][0];
            
          // $modelUserLogic = Loader::model('User','logic');//可能是链接数据库
          // $ret=$modelUserLogic->add($user);
          // print_r($user);die;
              $success=Db::name('xuhao')->insert($user); //批量插入数据
            }
              return $this->success('导入成功','index/import');
              
            }else{
                  return  array("resultcode" => -5,"resultmsg" =>"文件不存在","data" =>null);
            }


    }

    // 获取字符串中数字
    public function findNum($str=''){
    $str=trim($str);
    if(empty($str)){return '';}
    $temp=array('1','2','3','4','5','6','7','8','9','0');
    $result='';
    for($i=0;$i<strlen($str);$i++){
        if(in_array($str[$i],$temp)){
            $result.=$str[$i];
        }
    }
    return $result;
    }
    // 获取星期与周次
    public function jieciNum($str=''){
     $str=trim($str);
    if(empty($str)){return '';}
    $temp=substr($str,0,strlen($str)-1);
    $result='';
    $result=explode("[",$temp);
    return $result;

    }
    //对数据进行处理 生成总无课表
    public function wukebiao(){

    	  Db::execute('TRUNCATE table wp_zhouyi');
    	$wp = Db::name('jiaoshi')->where('xiaoqu','金明校区')->select(); 
    	
    	foreach ($wp as $key => $value) {
    		if ($value['jieci']=='1-2') {
    			$biao[$value['xingqi']][$value['number']]['1-2']=$value['lou'];
    		}
    		if ( $value['jieci']=='1-3') {
    			$biao[$value['xingqi']][$value['number']]['1-2']=$value['lou'].'3';
    		}
    		if ($value['jieci']=='3-4') {
    			$biao[$value['xingqi']][$value['number']]['3-4']=$value['lou'];
    		}
    		if ($value['jieci']=='7-8') {
    			$biao[$value['xingqi']][$value['number']]['7-8']=$value['lou'];
    		}
    		if ($value['jieci']=='7-9') {
    			$biao[$value['xingqi']][$value['number']]['7-8']=$value['lou'].'3';
    		}
    		if ($value['jieci']=='9-10') {
    			$biao[$value['xingqi']][$value['number']]['9-10']=$value['lou'];
    		}
    		if ($value['jieci']=='11-13' or $value['jieci']=='11-12') {
    			$biao[$value['xingqi']][$value['number']]['11-13']=$value['lou'];
    		}

    		
    		// var_dump($value);
    	}
       // var_dump($wp);
       var_dump($biao);
       foreach ($biao as $key => $value) {
       	$wuke['xingqi']= $key;
       	  foreach ($value as $ke => $val) {
       	  	$wuke['xingqi']= $key;
       	  	$wuke['jiaoshi']= $ke;
       	  	$wuke['one'] = isset($val['1-2'])?$val['1-2']:'';
       	$wuke['two'] = isset($val['3-4'])?$val['3-4']:'';
       	$wuke['three'] = isset($val['7-8'])?$val['7-8']:'';
       	$wuke['four'] = isset($val['9-10'])?$val['9-10']:'';
       	$wuke['five'] = isset($val['11-13'])?$val['11-13']:'';
       	$success=Db::name('zhouyi')->insert($wuke);
       	
       	  }
       	// $wuke['one'] = isset($value['1-2'])?$value['1-2']:'';
       	// $wuke['two'] = isset($value['3-4'])?$value['3-4']:'';
       	// $wuke['three'] = isset($value['7-8'])?$value['7-8']:'';
       	// $wuke['four'] = isset($value['9-10'])?$value['9-10']:'';
       	// $wuke['five'] = isset($value['11-13'])?$value['11-13']:'';
       
       	//var_dump($wuke);die;
       	 //$success=Db::name('zhouyi')->insert($wuke);
        // var_dump($wuke);
       }
       
    }
    //为导出做提前准备
   public function daochu($number,$louhao){
    
    $zhouyi = Db::name('zhouyi')->where('xingqi',$number)->order('jiaoshi esc')->select(); 
    
      $xuhao = Db::name($louhao)->order('xuhao esc')->select(); 
      
     foreach ($xuhao as $key => $value) {
       $wp= Db::name('zhouyi')->where('xingqi',$number)->where('jiaoshi',$value['xuhao'])->order('jiaoshi esc')->select();
       // var_dump($wp);die;
        if (!$wp) {

        	$wanzheng[$key]=array('xingqi' => $number ,
          'jiaoshi' => $value['xuhao'],
          'one' =>  '' ,
          'two' =>  '' ,
          'three' =>  '' ,
          'four' =>  '' ,
          'five' =>  '' );
        }else{
        	$wanzheng[$key]=array(
        		'id' => $wp[0]['id'],
        		'xingqi' => $wp[0]['xingqi'] ,
          'jiaoshi' => $wp[0]['jiaoshi'],
          'one' =>   $wp[0]['one'] ,
          'two' =>   $wp[0]['two'] ,
          'three' =>  $wp[0]['three'] ,
          'four' =>   $wp[0]['four'] ,
          'five' =>   $wp[0]['five'] );

        }
     	# code...
     }
     
     return $wanzheng;

   }
   //导出不同校区的无课表
    public function outwuke(){
     $this->out('xuhao');//xuhao为三号楼教室号
    
    }
   //导出总的无课表Excel
    public function out($louhao)
    {
        
        //导出
        $path = dirname(__FILE__); //找到当前脚本所在路径

        vendor("PHPExcel.PHPExcel.PHPExcel");
        vendor("PHPExcel.PHPExcel.Writer.IWriter");
        vendor("PHPExcel.PHPExcel.Writer.Abstract");
        vendor("PHPExcel.PHPExcel.Writer.Excel5");
        vendor("PHPExcel.PHPExcel.Writer.Excel2007");
        vendor("PHPExcel.PHPExcel.IOFactory");
        $objPHPExcel = new \PHPExcel();
        $objWriter = new \PHPExcel_Writer_Excel5($objPHPExcel);
        $objWriter = new \PHPExcel_Writer_Excel2007($objPHPExcel);
        $objPHPExcel->getActiveSheet()->setTitle('周一');      //设置sheet的名称
        $objPHPExcel->setActiveSheetIndex(0); 

        // 实例化完了之后就先把数据库里面的数据查出来 按照实际有课教室进行
        // $sql = Db::name('zhouyi')->where('xingqi','一')->order('jiaoshi esc')->select();
        // $two = Db::name('zhouyi')->where('xingqi','二')->order('jiaoshi esc')->select();
        // $three = Db::name('zhouyi')->where('xingqi','三')->order('jiaoshi esc')->select();
        // $four = Db::name('zhouyi')->where('xingqi','四')->order('jiaoshi esc')->select();
        // $five = Db::name('zhouyi')->where('xingqi','五')->order('jiaoshi esc')->select();
        // $six = Db::name('zhouyi')->where('xingqi','六')->order('jiaoshi esc')->select();
        // $seven = Db::name('zhouyi')->where('xingqi','日')->order('jiaoshi esc')->select();
        // 需要提前设置好教室号，但会把没课教室也显示
        $sql = $this->daochu('一',$louhao);
        $two = $this->daochu('二',$louhao);
        $three = $this->daochu('三',$louhao);
        $four = $this->daochu('四',$louhao);
        $five = $this->daochu('五',$louhao);
        $six = $this->daochu('六',$louhao);
        $seven = $this->daochu('日',$louhao);
       
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($sql);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $sql[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $sql[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $sql[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $sql[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $sql[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $sql[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        
        //周二
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(1);
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(1)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($two);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $two[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $two[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $two[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $two[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $two[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $two[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        $objPHPExcel->getActiveSheet()->setTitle('周二');      //设置sheet的名称
      
        //
        //周三
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(2);
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(2)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($three);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $three[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $three[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $three[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $three[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $three[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $three[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        $objPHPExcel->getActiveSheet()->setTitle('周三');      //设置sheet的名称
      
        //
         //周四
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(3);
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(3)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($four);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $four[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $four[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $four[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $four[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $four[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $four[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        $objPHPExcel->getActiveSheet()->setTitle('周四');      //设置sheet的名称
      
        //
        //         //周五
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(4);
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(4)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($five);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $five[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $five[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $five[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $five[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $five[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $five[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        $objPHPExcel->getActiveSheet()->setTitle('周五');      //设置sheet的名称
      
        //
        //        //         //周六
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(5);
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(5)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($six);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $six[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $six[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $six[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $six[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $six[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $six[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        $objPHPExcel->getActiveSheet()->setTitle('周六');      //设置sheet的名称
      
        //
        //        //        //         //周日
        $objPHPExcel->createSheet();
        $objPHPExcel->setActiveSheetIndex(6);
        // 设置表头信息
        $objPHPExcel->setActiveSheetIndex(6)
        ->setCellValue('A1', '教室')
        ->setCellValue('B1', '1-2节')
        ->setCellValue('C1', '3-4节')
        ->setCellValue('D1', '7-8节')
        ->setCellValue('E1', '9-10节')
        ->setCellValue('F1', '11-12节');

        /*--------------开始从数据库提取信息插入Excel表中------------------*/

        $i=2;  //定义一个i变量，目的是在循环输出数据是控制行数
        $count = count($seven);  //计算有多少条数据
        for ($i = 2; $i <= $count+1; $i++) {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $seven[$i-2]['jiaoshi']);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $i, $seven[$i-2]['one']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $i, $seven[$i-2]['two']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $i, $seven[$i-2]['three']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $i, $seven[$i-2]['four']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $i, $seven[$i-2]['five']);
        }

        
        /*--------------下面是设置其他信息------------------*/

        $objPHPExcel->getActiveSheet()->setTitle('周日');      //设置sheet的名称
      
        //

                          //设置sheet的起始位置
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   //通过PHPExcel_IOFactory的写函数将上面数据写出来
        
        $PHPWriter = \PHPExcel_IOFactory::createWriter( $objPHPExcel,"Excel2007");
            
        header('Content-Disposition: attachment;filename="设备列表.xlsx"');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
        exit; 
    }


}
