<?php
    $type = $_GET['type'];
    $act = $_GET["act"];
    $id = $_GET["id"];
    $params = $type." ".$act." ".$id;

    $url = null;
    if ($act == '02')
        $url = "Location:https://www.google.com/search?q=2";
    elseif($act == '03')
        $url = "Location:https://www.google.com/search?q=3";
    else
        $url = "Location:https://www.google.com/search?q=exit";

    header($url);

    if ($id != null){
        $path="python3 insert_log.py "; //需要注意的是：末尾要加一個空格
        passthru($path.$params); //等同於命令`python python.py 引數`，並接收列印出來的資訊 
    }

    exit;
?>
