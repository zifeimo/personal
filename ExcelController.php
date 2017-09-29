public function actionPage{
  header("Access-Control-Allow-Origen:*");
  $model = new Class();
  if(!empty($_POST)){
    $model = new Class();
    $model = load($_POST);
    if(\Yii::$app->request->isAjax){
      \Yii::$app->response->format = \yii\web\Response::FORMAT_JSON;
      return \yii\bootstrap\ActiveForm::validate($model);
    }
    $model->load(\Yii::$app->request->post());
    $name = $_FILES['Class']['name']['name'];
    $tag_data = ExcelImport::SaveFile($model,$name);
    //循环每条数据验证与存储
    ExcelImport::EachImport($tag_data);
    return $this->render('template',[
    ]);
  }
  return $this->render('excel',[
            'model' => $model,
        ]);
}
