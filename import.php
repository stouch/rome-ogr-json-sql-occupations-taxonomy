
/******** INPUT VARIABLES ********/

// Originally created using https://github.com/PHPOffice/PHPExcel
require_once '/path/to/PHPExcel/Classes/PHPExcel/IOFactory.php';

// xlsx from https://www.data.gouv.fr/en/datasets/repertoire-operationnel-des-metiers-et-des-emplois-rome/
$tmpfname = "/path/to/ROME_ArboPrincipale.xlsx";

// Your JSON and SQL outputs
$outputJSONPath = "/path/to/ROME_ArboPrincipale.json";
$outputSQLPath = "/path/to/ROME_ArboPrincipale.sql";

/****** END INPUT VARIABLES ******/


$outputJSON = [];
$outputSQL = '

CREATE TABLE rome_taxonomy_category (
    rome_prefix varchar(10) PRIMARY KEY,
    parent_rome_prefix varchar(10),
    label varchar(255) NOT NULL
);
CREATE TABLE rome_taxonomy_ogr (
    ogr int PRIMARY KEY,
    label varchar(255) NOT NULL,
    lvl_1_rome_prefix varchar(10) NOT NULL,
    lvl_2_rome_prefix varchar(10) NOT NULL,
    lvl_3_rome varchar(10) NOT NULL
);
';

$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
$excelObj = $excelReader->load($tmpfname);

// ROME codes are in the second sheet :
$sheets = $excelObj->getAllSheets();
$worksheet = $excelObj->getSheet(1);
$lastRow = $worksheet->getHighestRow();

$ROMECategory = [];
for ($row = 2; $row <= $lastRow; $row++) {
    $A = trim($worksheet->getCell('A' . $row)->getValue(), " ");
    if (empty($A)) {
        continue;
    }
    $B = trim($worksheet->getCell('B' . $row)->getValue(), " ");
    $C = trim($worksheet->getCell('C' . $row)->getValue(), " ");
    $D = trim($worksheet->getCell('D' . $row)->getValue(), " ");
    $E = trim($worksheet->getCell('E' . $row)->getValue(), " ");
    if ($A && $B && $C && !empty($E)) {
        // OGR occupation
        $ROMECategory[$A][1][$B][1][$C][1][$E] = $D;
    } else {
        // ROME Category
        if ($C != "") {
            $ROMECategory[$A][1][$B][1][$C] = [
                $D,
                []
            ];
        } else if ($B != "") {
            $ROMECategory[$A][1][$B] = [
                $D,
                []
            ];
        } else {
            $ROMECategory[$A] = [
                $D,
                []
            ];
        }
    }
}

foreach ($ROMECategory as $catLvl1 => $lvl1) {
    $outputSQL .= 'INSERT INTO rome_taxonomy_category (rome_prefix, parent_rome_prefix, label) VALUES (' .
        '"' . $catLvl1 . '",' .
        'NULL,' .
        '"' . $lvl1[0] . '");' . "\n";
    foreach ($lvl1[1] as $catLvl2 => $lvl2) {
        $outputSQL .= 'INSERT INTO rome_taxonomy_category (rome_prefix, parent_rome_prefix, label) VALUES (' .
            '"' . $catLvl1 . $catLvl2 . '",' .
            '"' . $catLvl1 . '",' .
            '"' . $lvl2[0] . '");' . "\n";
        foreach ($lvl2[1] as $catLvl3 => $lvl3) {
            $outputSQL .= 'INSERT INTO rome_taxonomy_category (rome_prefix, parent_rome_prefix, label) VALUES (' .
                '"' . $catLvl1 . $catLvl2 . $catLvl3 . '",' .
                '"' . $catLvl1 . $catLvl2 . '",' .
                '"' . $lvl3[0] . '");' . "\n";
            foreach ($lvl3[1] as $ogr => $occupation) {
                $outputSQL .= 'INSERT INTO rome_taxonomy_ogr (ogr, label, lvl_1_rome_prefix, lvl_2_rome_prefix, lvl_3_rome) VALUES (' .
                    $ogr . ',' .
                    '"' . $occupation . '",' .
                    '"' . $catLvl1 . '",' .
                    '"' . $catLvl1 . $catLvl2 . '",' .
                    '"' . $catLvl1 . $catLvl2 . $catLvl3 . '");' . "\n";
                $outputJSON[] = [
                    'ogr' => $ogr,
                    'label' => $occupation,
                    'lvl_1_rome_prefix' => $catLvl1,
                    'lvl_2_rome_prefix' => $catLvl2,
                    'lvl_3_rome' => $catLvl3,
                ];
            }
        }
    }
}

file_put_contents($outputJSONPath, json_encode($outputJSON, JSON_PRETTY_PRINT));
file_put_contents($outputSQLPath, $outputSQL);
