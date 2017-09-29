<?php

use yii\helpers\Html;
use kartik\widgets\ActiveForm;
use kartik\builder\Form;
use kartik\select2\Select2;
use kartik\widgets\FileInput;

?>
<div>

    <?php $form = ActiveForm::begin(
        [
            //'action' => ['/excel/import'],
            'type' => ActiveForm::TYPE_HORIZONTAL,
            'id' => 'form-AddStudent',
            'enableAjaxValidation' => true,
            'enableClientValidation' => true,
            'options' => ['enctype' => 'multipart/form-data']]) ?>
    <div class="panel panel-info">
        <div class="panel-heading">上传Excel</div>
        <div class="panel-body">
            <?= $form->field($model, 'user',['labelOptions'=>['label'=>'Excel表']])->widget(FileInput::classname(), [
                'options' => [
                    'accept' => 'application/*',
                    'multiple' => true,
                ],]);
            ?>
        </div>
    </div>

    <div class="form-group">
        <?= Html::submitButton($model->isNewRecord ? '创建' : '更新', ['class' => $model->isNewRecord ? 'btn btn-success pull-right' : 'btn btn-primary pull-right']) ?>
    </div>

    <?php ActiveForm::end(); ?>

</div>
