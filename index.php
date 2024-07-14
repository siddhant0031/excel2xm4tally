<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <title>Upload file</title>
</head>
<style media="screen">
  *{
    padding: 0;
    margin: 0;
  }
  div.main_wrapper{
    position: relative;
    width: 100vw;
    height: 100vh;
    padding: 10px;
    box-sizing: border-box;
  }
  div.title_wrapper{
    width: 100%;
    text-align: center;
    font-size: 3vh;

  }
  div.title_wrapper span:first-child{
    font-size: 5vh;
  }
  div.wrapper {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    margin-top: 5vh;
  }
  div.link_wrapper{
    text-align: center;
    width: 80%;
    max-width: 800px;
    padding: 10px;
    box-sizing: border-box;

  }
  div.link_wrapper a{
    text-decoration: none;
    color: white;
    background: blue;
    padding:5px;
    box-sizing: border-box;
    border-radius: 5px;
    cursor: pointer;
  }
  div.form_wrapper{
    width: 90%;
    max-width: 600px;
    height: 100px;
    border-radius: 10px;
    border: 1px solid blue;
    padding: 10px;
    box-sizing: border-box;
  }
  div.form_wrapper form{
    display: flex;
    justify-content: space-around;
    height: 100%;
    align-items:  center;
  }
  div.form_wrapper button{
    padding: 10px;
    color:white;
    outline: none;
    border:none;
    background: blue;
    border-radius: 5px;
    cursor: pointer;

  }
  div.form_wrapper form label{
    color: blue;
    font-size: 20px;
    margin-bottom: 10px;
  }
  div.form_wrapper form input[type='file']{
    cursor: pointer;
  }
</style>
<body>
<div class="main_wrapper">
  <div class="title_wrapper">
    <span class="title">Mini Project</span><br>
    <span class="title">Excel to XML converter for Tally</span>
  </div>
  <div class="wrapper">
    <div class="link_wrapper">
      <span>Please Click Next Button For Download Voucher Template</span>
      <a href="download_template.php?filename=template4voucher.xlsx">Template</a>
    </div>
    <div class="form_wrapper">
      <form action="<?php $_SERVER['PHP_SELF']; ?>" method="post" enctype="multipart/form-data">
        <div >
          <label for="uploaded_file">Upload The Template</label><br>
          <input type="file" id="uploaded_file" name="file" required>
        </div>
        <div >
          <button type="submit" name="button">Convert & Download</button>
        </div>
      </form>
    </div>
  </div>
</div>
  <div>
        <?php

        include 'vendor/autoload.php';
        use PhpOffice\PhpSpreadsheet\IOFactory;
        use PhpOffice\PhpSpreadsheet\spreadsheet;
        if (isset($_FILES['file'])) {
          $fileName =explode('.',$_FILES['file']['name']);
          $extension = ['xlsx','xls'];
          $xtention= end($fileName);
          $xmlFileName = "Excel2".explode('.',basename($_FILES['file']['tmp_name']))[0];
          if (in_array($xtention,$extension)) {

          // echo readfile($_FILES['file']['tmp_name']);

        // echo readfile("read me.txt");
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
        $reader->setReadDataOnly(TRUE);
        $spreadsheet = $reader->load($_FILES['file']['tmp_name']);

        $worksheet = $spreadsheet->getSheet("0");
        // code of xml that are constant
         $data = '<? xml version = "1.0" encoding = "UTF-8" ?>

            <ENVELOPE>
                <HEADER>
                    <TALLYREQUEST>Import Data</TALLYREQUEST>
                </HEADER>


                <BODY>
                <IMPORTDATA>
                    <REQUESTDESC>

                        <REPORTNAME>All Masters</REPORTNAME>
                        <STATICVARIABLES>
                            <SVCURRENTCOMPANY>Siddhant Sanmane</SVCURRENTCOMPANY>
                        </STATICVARIABLES>
                    </REQUESTDESC>

                    <REQUESTDATA>';
        // end of starting lines of xml format that are constnat
        $row1=0;

        foreach ($worksheet->getRowIterator() as $row) {
          $date;
          $narration="";
          $voucher = "";
          $LedgerName="";
          $PartyName="";
          $ledgerDetails ="";
          if ($row1 == 0) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false);
            $i = 0;
            foreach ($cellIterator as $cell) {
              $even = ['Date','Voucher','Ledger Name'];
              $odd = ['Narration','Voucher Number','Ledger Amt'];
              if ($i%2 == 0 && !in_array($cell->getValue(),$even)) {
                die("error 1: use Proper Template");
              }elseif ($i%2 != 0 && !in_array($cell->getValue(),$odd)) {
                die("error 2: use Proper Template");
              }
              $i++;
            }
          }elseif ($row1 != 0) {
            // code...
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false); // This loops through all cells,
                                                               //    even if a cell value is not set.
                                                               // For 'TRUE', we loop through cells
                                                               //    only when their value is set.
                                                               // If this method is not called,
                                                               //    the default value is 'false'.

            $i = 0;
            foreach ($cellIterator as $cell) {
              if ($i == 0) {
                $date =   date("Ymd",($cell->getValue()-25569)*86400) ;
              }elseif ($i == 1) {
                // code...
                  $narration =   $cell->getValue() ;

              }elseif ($i == 2) {
                  $voucher =  $cell->getValue() ;
              }elseif ($i == 3) {
                  $voucherNumber = $cell->getValue();
              }elseif ($i%2 == 0) {
                $LedgerName = $cell->getValue();
                $ledgerDetails .= "<ALLLEDGERENTRIES.LIST>";
                  $ledgerDetails .="<LEDGERNAME>" . $LedgerName ."</LEDGERNAME>" ;
                  $ledgerDetails .= "<REMOVEZEROENTRIES>NO</REMOVEZEROENTRIES>
                            <LEDGERFROMITEM>NO</LEDGERFROMITEM>";
              }elseif($i%2 !=0) {
                  $ledgerAmt =   $cell->getValue() ;
                  if ($ledgerAmt < 0) {
                    $deemToPositive = "YES";
                  }else {
                    $deemToPositive = "NO";
                  }
                  $ledgerDetails .= "<ISDEEMEDPOSITIVE>".$deemToPositive."</ISDEEMEDPOSITIVE>
                  <AMOUNT>".$ledgerAmt."</AMOUNT>
                </ALLLEDGERENTRIES.LIST>";
              }
              $i++;
            }
            //code for gathering all the data after each row

            $data .= '<TALLYMESSAGE xmlns:UDF="TallyUDF">
              <VOUCHER ACTION="Create" VCHTYPE="'.$voucher.' ">
              <VOUCHERTYPENAME>'.$voucher.'</VOUCHERTYPENAME>
                <DATE>'.$date.'</DATE>
                <VOUCHERNUMBER>'.$voucherNumber.'</VOUCHERNUMBER>
                <PARTYLEDGERNAME>'.$PartyName.'</PARTYLEDGERNAME>
                <NARRATION>'.$narration.'</NARRATION>
                <EFFECTIVEDATE>'.$date.'</EFFECTIVEDATE>'.$ledgerDetails.'</VOUCHER> </TALLYMESSAGE>';
              // end -- code for gathering all the data after each row

          }
          $row1++;

        }

        //ending lines of constant xml format
        $data .="</REQUESTDATA>
    </IMPORTDATA>
    </BODY>
    </ENVELOPE>";
    // echo $data;

        //end of ending lines of constant xml format

        // code for creating and writing the data into xml file

        $newFile = fopen("xml/{$xmlFileName}.xml","w");
        fwrite($newFile,$data);
        fclose($newFile);
        header("Location: download.php?filename={$xmlFileName}.xml");
      }else{
        echo "<div>use only excel file with extention 'xlsx'or'xls'</div>";
      }
      }

         ?>
  </div>
</body>
</html>
