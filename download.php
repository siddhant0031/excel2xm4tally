<?php
if (isset($_GET['filename'])) {
  $fileName = $_GET['filename'];
  $filePath = 'xml/'.$fileName;
  if (file_exists($filePath)) {
    header("Cache-Control: public");
    header("Content-Description:File Transfer");
    header("Content-Disposition:attachment; filename=$fileName");
    header("Content-Type: applicaiton/zip");
    header("Content-Transfer-Encoding:binary");
    readfile($filePath);
    unlink($filePath);
    exit();
  }else{
    echo "<div>Please Try Again</div>";
  }

}
    ?>
