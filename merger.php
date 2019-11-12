<?php

/*
find all csv from the same pattern
merge them into 1 excel file, with each pair on a separate sheet
combined results sheet has the combined settings from all pairs, and filters the common settins
    that exist on all pairs

ultima test: test the filtered settings as far back as you can till sept 2019, and plug them in 
    to the excel file to filter the best settings possible
*/


require_once(__DIR__ . "/vendor/autoload.php");

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

function print_r2($a, $out=false)
{
    $outStr = "";
    if ($out)
    {
        $outStr =  "<pre>" . print_r($a, true) . "</pre>";
        return $outStr;
    }

    echo "<pre>";
    echo print_r($a, true);
    echo "</pre>";
}

function pt(string $s)
{
    echo "[" . (new DateTime())->format("Y-m-d H:i:s") . "] ".$s."\n";
}

class Merger
{
    private static $configPath = __DIR__."/config.json";
    private static $config = [];
    private static $configTemplate = 
    [
        "fileFormatRegex" => "\[(.*?)\]_(\w{6})_(\d{4}\.\d{2}\.\d{2})_(\d{4}\.\d{2}\.\d{2})-(\d{14})",
        "resultsFolder" => "results",
        "mergeAllOrNew" => false,
    ];
    private static $resultsFolder = __DIR__."/results/";

    public static function init()
    {
        register_shutdown_function(function()
        {
            $a = error_get_last();
            if ($a != null)
                pt(print_r2($a, true));
        });

        self::loadConfig();

        self::processCmdArgs();

        Cell\Cell::setValueBinder( new Cell\AdvancedValueBinder() );

        /*
        $pool = new \Cache\Adapter\Apcu\ApcuCachePool();
        $simpleCache = new \Cache\Bridge\SimpleCache\SimpleCacheBridge($pool);

        \PhpOffice\PhpSpreadsheet\Settings::setCache($simpleCache);
        */

        self::startProcess();
    }

    public static function processCmdArgs()
    {
        global $argc, $argv;

        if ($argc > 2)
        {
            $args = array_splice($argv, 1);
            foreach ($args as $key => $value)
            {
                switch ($value)
                {
                    case 'config':
                        return;
                        break;

                    case 'results':
                        self::$config["resultsFolder"] = $args[$key + 1];
                        return;
                        break;

                    default:
                        echo "\nRun with the config.json data: config";
                        echo "\nRun with a custom results folder: results path_to_folder";
                        die();
                        break;
                }
            }

        }
    }

    public static function startProcess()
    {
        $grouped = self::getGroupedFiles();

        pt("To merge ".count($grouped).": \n-".implode("\n-", array_keys($grouped))."\n");

        //create excel files for each group
        foreach ($grouped as $key => $value)
        {
             self::createExcelForGroup($key, $value);
        }

        pt("End!");
    }

    public static function createExcelForGroup(string $indicator, array $groupData)
    {
        $arr = [];
        foreach (array_column($groupData, "matches") as $key => $value)
        {
            $arr[] = $value[2];
        }
        $arr = array_unique($arr);
        pt("Creating excel for '$indicator' (".count($arr).") ".implode(",", $arr));

        $splitIndiname = explode("|", $indicator);
        $getStyle = function(string $col, bool $combinedSheet = false)
        {
            $cols = $combinedSheet ? ["A", "B", "C", "D", "G"] : ["C", "D", "E", "F", "I"];

            $t = 
            [
                'fill' =>
                [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                    'rotation' => 90,
                    'startColor' =>
                    [
                        'argb' => 'FFFFEF9C', // yellow
                    ],
                    'endColor' =>
                    [
                        'argb' => 'FF63BE7B', // green
                    ],
                ],
            ];

            if (!in_array(strtoupper($col), $cols))
            {
                $t["fill"]["startColor"]["argb"] = 'FF63BE7B';
                $t["fill"]["endColor"]["argb"] = 'FFFFEF9C';
            }

            return $t;
        };

        $indexesCondition = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
        $indexesCondition->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
        $indexesCondition->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHAN);
        $indexesCondition->addCondition('0');
        $indexesCondition->getStyle()
            ->applyFromArray(
            [
                'fill' =>
                [
                    'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                    'color' =>
                    [
                        'argb' => 'FFC6EFCE',
                    ],
                ],
                'font' =>
                [
                    'color' =>
                    [
                        'arbg' => 'FF006100',
                    ]
                ],
            ]);

        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->setTitle("Combined Results");

        $uniqueInputs = [];
        $headers = ["Inputs", "Nº Idxs"];
        $headersSet = false;

        // each currency pair
        $pairsAdded = [];
        foreach ($groupData as $key => $value)
        {
            $file = $value["file"];
            $matches = $value["matches"];

            $fileData = file_get_contents($file);

            $lines = explode("\n", $fileData);
            if (count($lines) === 1)
            {
                pt("File '".pathinfo($file, PATHINFO_FILENAME)."' doesn't have data.");
                continue;
                //return;
            }

            if (in_array($matches[2], $pairsAdded))
                continue;

            if (!$headersSet)
            {
                $headersSet = true;

                $h = explode(";", $lines[0]);
                $h = array_splice($h, 1);
                $headers = array_merge($h, $headers);
            }
            
            $dataInArray = [];
            $range = null;
            $countRange = 0;
            $inputsFunc = [];
            $inputsKey = 0;
            foreach ($lines as $key => $value)
            {
                //$value = str_replace(".", ",", $value);
                $lineData = explode(";", $value);

                if (strpos($lineData[7], "100") !== false
                //|| strpos($lineData[7], "999") !== false 
                || strpos($lineData[7], ";0.000") !== false)
                {
                    continue;
                }

                $lineData[7] = $lineData[7] . "%";
                /*
                $lineData[6] = explode(".", str_replace("%", "", $lineData[6]));
                $lineData[6] = "0.".$lineData[6][0].$lineData[6][1];

                $lineData[7] = explode(".", $lineData[7]);
                $lineData[7] = "0.".$lineData[7][0].$lineData[7][1];
                */

                if ($key === 0)
                {
                    $dataInArray[] = array_merge(["Inputs"], $lineData);
                    continue;
                }

                $inputsKey++;

                $dataInsert = array_merge(
                    [
                        ""//$inputsFunc
                    ],
                    $lineData
                );

                if ($range === null)
                {
                    $a = Coordinate::columnIndexFromString("J");
                    $b = Coordinate::columnIndexFromString(Coordinate::stringFromColumnIndex(count($dataInsert)));
                    $range = range($a, $b);
                    $countRange = count($range);

                    for ($i=1; $i <= $countRange; $i++)
                    {
                        $v = Coordinate::stringFromColumnIndex($range[$i - 1]);
                        $inputsFunc[] =
                            "IF(".
                                "OR(".
                                    "IFERROR(FIND(\"INPUT_\",".$v."__k__),0),".
                                    "IFERROR(FIND(\"Input\",".$v."__k__),0),".
                                    "IFERROR(FIND(\"Indicator\",".$v."__k__),0)".
                                "),".
                            $v."__k__,\"\")";
                    }
                    $inputsFunc = "=CONCATENATE(" . implode(",", $inputsFunc) . ")";
                }

                $inputsFuncTemp = preg_replace("/__k__/", $inputsKey + 1, $inputsFunc);
                /*
                if ($key === 1)
                {
                    echo "$inputsFuncTemp";
                    die();
                }
                */

                $dataInsert[0] = $inputsFuncTemp;

                $i = $key + 1;
                $dataInArray[] = $dataInsert;

                //print_r2($dataInArray);die();
            }

            if (count($dataInArray) <= 1)
                continue;

            $newWorksheet = new Worksheet($spreadsheet, $matches[2]);
            $newWorksheet->setCellValue("A1", "Inputs");
            $spreadsheet->addSheet($newWorksheet);

            $allRange = $newWorksheet->calculateWorksheetDimension();
            $dataRange = "A2:".$newWorksheet->getHighestColumn() . $newWorksheet->getHighestRow();
            $newWorksheet->fromArray($dataInArray, null, "A1")
                ->setAutoFilter("A1:".$newWorksheet->getHighestColumn().$newWorksheet->getHighestRow()) // filter columns
                ->freezePane("A2"); // freeze pane

            /*
            foreach (range("A", $newWorksheet->getHighestColumn()) as $key => $value)
            {
                $newWorksheet->getColumnDimension($value)->setAutoSize(true); // columns auto sized
            }

            $newWorksheet->getStyle('A1:' . $newWorksheet->getHighestColumn() . "1")
                    ->applyFromArray(
                    [
                        'fill' =>
                        [
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'color' =>
                            [
                                'argb' => 'FFFCE5CD', // yellow
                            ],
                        ],
                    ]);

            // apply gradients
            foreach (["C", "D", "E", "F", "I"] as $value)
            {
                $newWorksheet->getStyle($value . '2:' . $value . $newWorksheet->getHighestRow())
                    ->applyFromArray($getStyle($value, false));
            }
            */

            /*
            $a = $newWorksheet->rangeToArray("A2:A".$newWorksheet->getHighestRow());
            $sheetData = $newWorksheet->toArray(null, false, false, true);
            print_r2($a);
            die();
            */

            // calculated values
            $uniqueInputs = array_merge($uniqueInputs, $newWorksheet->rangeToArray("A2:A".$newWorksheet->getHighestRow()));
            $headers[] = $pairsAdded[] = $matches[2];
        }

        if (count($uniqueInputs) === 0)
        {
            pt("Skipping '".pathinfo($file, PATHINFO_FILENAME)."' because it doesn't have data to merge.");
            return;
        }
        
        $uniqueInputs = array_values(array_filter(array_unique(call_user_func_array('array_merge', $uniqueInputs))));

        // update combined results sheet
        $sheet = $spreadsheet->getSheet(0);
        $sheet->fromArray(
            [
                ["C", "D", "E", "F", "I", "H", "I"],
                $headers,
            ])
            ->setAutoFilter($allRange); // filter columns

        /*
        $sheetData = $sheet->toArray(null, false, false, true);
        print_r2($sheetData);
        die();
        */

        $combinedResultsData = [];
        $sc = $sheet->getHighestDataColumn(2);
        $hc = "J";
        $highestCol = Coordinate::columnIndexFromString($sc);
        $indexesStartCol = Coordinate::columnIndexFromString($hc);
        $range = range($indexesStartCol, $highestCol);
        foreach ($uniqueInputs as $key => $value)
        {
            $i = $key + 3;

            $temp =
            [
                '=SUM(', 
                '=SUM(', 
                '=SUM(', 
                '=SUM(', 
                '=SUM(', 
                '=SUM(', 
                '=SUM(', 
                $value,
                '=COUNT('.$sc.$i.':'.$hc.$i.')',
            ];
            $c = count($range);
            for ($j=0; $j < $c; $j++)
            {
                $isLast = $j + 1 >= $c;
                $l = Coordinate::stringFromColumnIndex($range[$j]);

                $temp[0] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&A$1&$'.$l.$i.'' . ( $isLast ? '), 0))' : "), 0)," );
                $temp[1] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&B$1&$'.$l.$i.'' . ( $isLast ? '), 0))' : "), 0)," );
                $temp[2] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&C$1&$'.$l.$i.'' . ( $isLast ? '), 0)) / $I'.$i."" : "), 0)," );
                $temp[3] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&D$1&$'.$l.$i.'' . ( $isLast ? '), 0)) / $I'.$i."" : "), 0)," );
                $temp[4] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&E$1&$'.$l.$i.'' . ( $isLast ? '), 0))' : "), 0)," );
                $temp[5] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&F$1&$'.$l.$i.'' . ( $isLast ? '), 0)) / $I'.$i."" : "), 0)," );
                $temp[6] .= 'IFERROR(INDIRECT($'.$l.'$2&"!"&G$1&$'.$l.$i.'' . ( $isLast ? '), 0)) / $I'.$i."" : "), 0)," );

                $temp[] = '=MATCH(TRUE,INDEX(INDIRECT('.$l.'$2&"!$A:$A")=H'.$i.',0),0)';
            }

            $combinedResultsData[] = $temp;

            if ($key === 0)
            {
                //print_r2($combinedResultsData);die();
            }
            //break;
        }

        $sheet->fromArray($combinedResultsData, null, "A3")
            ->freezePane("A3")
            ->setAutoFilter("A2:".$sheet->getHighestDataColumn().$sheet->getHighestRow());


        foreach (range("A", "G") as $key => $value)
        {
            /*
            if (!in_array($value, ["A", "E", "G"]))
                continue;
            */

            for ($i=3; $i <= $sheet->getHighestRow(); $i++)
            {
                if (in_array($value, ["F", "G"])) // Drawdown %, OnTester
                {
                    $sheet->getCell($value.$i)->getStyle()
                        ->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_PERCENTAGE);
                }
                else if (in_array($value, ["A", "E"])) // Profit, Drawdown $
                {
                    $sheet->getCell($value.$i)->getStyle()
                        ->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
                }
                else if (in_array($value, ["C", "D"])) // Profit Factor, Expected Payoff
                {
                    $sheet->getCell($value.$i)->getStyle()
                        ->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00);
                }
            }
        }

        /*
        $sheetData = $sheet->toArray(null, false, false, true);
        print_r2($sheetData);
        die();
        */

        foreach (range("A", $sheet->getHighestDataColumn(2)) as $key => $value)
        {
            $sheet->getColumnDimension($value)->setAutoSize(true); // columns auto sized
        }

        // index colors
        $sheet->getStyle($indexesStartCol."3:".$highestCol.$sheet->getHighestRow())->setConditionalStyles([$indexesCondition]);

        $writer = new Xlsx($spreadsheet);

        if (!file_exists(self::$resultsFolder))
            mkdir(self::$resultsFolder);
        $savedFilename = "merged_[".$splitIndiname[0]."]_".$splitIndiname[1]."_".$splitIndiname[2]."_".count($pairsAdded)."pairs.xlsx";
        $saved = $writer->save(self::$resultsFolder.$savedFilename);

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);

        pt("## Merged: $savedFilename");
        //die();
    }

    // grab all results files and group them by date and indicator
    public static function getGroupedFiles()
    {
        $folder = self::$config["resultsFolder"];
        $files = glob($folder."/*.csv");

        $grouped = [];
        foreach ($files as $key => $f)
        {
            $filename = pathinfo($f, PATHINFO_FILENAME);

            if (!preg_match_all("/".self::$config["fileFormatRegex"]."/", $filename, $g))
            {
                echo "\nCan't match regex with file '$filename'";
                continue;
            }

            $g = call_user_func_array('array_merge', $g);

            $indi = $g[1] . "|" . $g[3] . "|" . $g[4];

            if (!array_key_exists($indi, $grouped))
                $grouped[$indi] = [];

            $grouped[$indi][] = 
            [
                "file" => $f,
                "matches" => $g,
            ];
        }

        if (count($grouped) === 0)
        {
            echo "\nNo results found. Is the \"resultsFolder\" config field correct?";
            die();
        }

        // check those already merged
        if (!self::$config["mergeAllOrNew"])
        {
            foreach ($grouped as $key => $value)
            {
                $splitIndiname = explode("|", $key);

                $savedFilename = "merged_[".$splitIndiname[0]."]_".$splitIndiname[1]."_".$splitIndiname[2]."_*";//.count($value)."pairs.xlsx";
                $files = glob(self::$resultsFolder . $savedFilename);
                if(count($files) > 0)
                    unset($grouped[$key]);

                /*
                $savedFilename = "merged_[".$splitIndiname[0]."]_".$splitIndiname[1]."_".$splitIndiname[2]."_".count($value)."pairs.xlsx";
                if (file_exists(self::$resultsFolder . $savedFilename))
                {
                    unset($grouped[$key]);
                }
                */
            }
        }

        return $grouped;
    }

    // loads config.json file
    private static function loadConfig(bool $reset = false)
    {
        if (!file_exists(self::$configPath) || $reset)
        {
            file_put_contents(self::$configPath, json_encode(self::$configTemplate), JSON_PRETTY_PRINT);

            echo("\nConfig file generated (config.json). Please change the following items if needed:\n- 'fileFormatRegex': results file format in regex\n");
            die();
        }

        self::$config = json_decode(file_get_contents(self::$configPath), true);
    }

    // save in memory config to file
    private static function saveConfig()
    {
        file_put_contents(self::$configPath, json_encode(self::$config, JSON_PRETTY_PRINT));
    }
}

if (php_sapi_name() === "cli")
    Merger::init();