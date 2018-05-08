<?php

    ini_set('display_errors', 1);
    error_reporting(E_ALL);

   if(!@copy('http://www.org.***/***_opt.zip','./test.zip'))
    {
        $errors= error_get_last();
        echo "COPY ERROR: ".$errors['type'];
        echo "<br />\n".$errors['message'];
    } else {
     echo "File copied from remote!\n";
    }

   $zip = new ZipArchive;
    $file = realpath("test.zip");
    $res = $zip->open($file);
    if ($res === TRUE) {
        $zip->extractTo('unzip');
        $zip->close();
        echo "unzip ok\n";
    } else {
        echo "failed, code:\n" . $res;
    }


require_once 'PHPExcel.php';
$pExcel = PHPExcel_IOFactory::load('unzip/'.'***_opt.xls');

// Цикл по листам Excel-файла
foreach ($pExcel->getWorksheetIterator() as $worksheet) {
    $tables[] = $worksheet->toArray();
}
$host = '192.168.*.***';
$dbname = 'sk******';
$options = [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION];
$username = 'v*******';
$passwd = '*****N******';

$DBH = new PDO("dblib:host=" . $host . ";dbname=" . $dbname, $username, $passwd, $options);

$zapros = "SELECT dbo.k******()";
$kurs = $DBH->query($zapros);

$kursResult;
foreach ($kurs as $row) {
    ($kursResult = $row[0]);
}

function searchArtikul($artikul, $brand) {
    $brand = trim($brand);
    $artikul = trim($artikul);

    $brandLocating = stripos($artikul, $brand);
    $brendLength = strlen($brand);
    $resultString = substr($artikul, $brandLocating + $brendLength);
    if (($brandLocating !== false) && (($brandLocating + $brendLength) != strlen($artikul))) {
        $resultString = pruningString($resultString);
        return $resultString;
    } else {
        return $artikul;
    }
}

function pruningString($resultString) {
    $firstCharacter = substr($resultString, 0, 1);
    $checkingFirstCharacter = preg_match("|^[a-zа-яA-ZА-ЯёЁ0-9]|U", $firstCharacter);
    while ($checkingFirstCharacter == '0') {
        $resultString = trim(substr($resultString, 1));
        $firstCharacter = substr($resultString, 0, 1);
        $checkingFirstCharacter = preg_match("|^[a-zа-яA-ZА-ЯёЁ0-9]|U", $firstCharacter);
    }
    return $resultString;
}

//Подготовленный запрос
$sql = "INSERT INTO [s*******].[dbo].[X********] ([klientID],[Artikul],[price1],[price_UAH],[nal],[PosCode],[Brand], [priceRRP_UAH])
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
$result = $DBH->prepare($sql);

foreach ($tables as $table) {
    foreach ($table as $row) {
        $price =  str_replace(',','', $row[4]);
        if (!is_numeric($price)) {
            continue;
        }
        $klientID = 5856;
        $artikul = iconv("UTF-8", "windows-1251", searchArtikul($row[2], $row[1]));
        $posCode =  str_replace(',','', $row[0] );
        $price1 = round($price / $kursResult, 2);
        $priceUAH = $price;
        if (!is_numeric(str_replace(',','', $row[5]))) {
            $priceRrpUah = 'null';
        } else {
            $priceRrpUah = str_replace(',','', $row[5]);
        }
        $nal = iconv("UTF-8", "windows-1251", $row[6]);
        $brand = iconv("UTF-8", "windows-1251", $row[1]);

        $result->execute(array($klientID, "$artikul", $price1, $priceUAH, "$nal", "$posCode", "$brand", $priceRrpUah));

    }

}
echo "Table created\n";
?>
