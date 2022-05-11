<?php
    
Class MagicRush
{
    public $start, $xls, $sheetOne, $sheetTwo, $objWriter;
    
    function __construct()
    {
        echo 'Создание ' . __CLASS__;
        echo '<br>';echo '<br>';
        $this->start = microtime(true);
        require_once __DIR__ . '/vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
        require_once __DIR__ . '/vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';

        $this->xls = PHPExcel_IOFactory::load('D:\MagicRushv2.xlsx');

        $this->sheetOne = $this->xls->setActiveSheetIndex(0);
        $this->sheetTwo = $this->xls->setActiveSheetIndex(1);
    }
    
    function __destruct()
    {
        $this->objWriter = PHPExcel_IOFactory::createWriter($this->xls, 'Excel2007');
        $this->objWriter->save('D:\MagicRushv2.xlsx');
        echo '<br>';echo '<br>';
        echo 'Уничтожение ' . __CLASS__;
        echo '<br>';echo '<br>';
        echo 'Время выполнения скрипта: '.round(microtime(true) - $this->start, 4).' сек.';
    }

    /**
	 *	Get a list of heroes
	 *
	 *	@return array       
	 */
    public function getHeroes():array
    {
        echo 'Инициализация геров:';echo '<br>';
        $heroes = array();
        for ($i = 3; ;$i++) 
        {   
            $CellRow = $this->sheetOne->getCell('A'.$i)->getValue();
            if (!empty($CellRow) && $CellRow != 'Итог:') {
                $heroes['A'.$i] = trim($this->sheetOne->getCell('A'.$i)->getValue());
            } else break;
        }
        echo 'Инициализация героев успешно завершено';echo '<br>';
        return $heroes;
    }
        
    /**
	 *	Search for the power of heroes by matching the name of the hero. 
     *  Calculation of the total power of the team of heroes. 
     *  Comparison of the current and recommended power of the heroes and changing the background of the cell. 
     *  Event "Battle of Heroes"
	 *	     
     *  @param array        $heroes
	 */
    public function calculationBattleOfHeroes(array $heroes)
    {
        echo 'Инициализация события Битва Героев:';echo '<br>';
        for ($i = 3;  ; $i++) 
        {
            $CellRow = $this->sheetTwo->getCell('A'.$i)->getValue();

            if (empty($CellRow)) {
                break;
            }

            unset($power);
            $power[5];

            //  Current total power
            for($j = 'A'; $j <= 'F'; $j++)
            {
                if (!empty($CellRow)) {
                    $nameHero = $this->sheetTwo->getCell($j.$i)->getValue();
                    $nameHero = trim($nameHero);
                    if (in_array($nameHero, $heroes)) {
                        preg_match('/\d+/', array_shift(array_keys($heroes, $nameHero)), $mathes);
                        $power[] = $this->sheetOne->getCell('C'.array_shift($mathes))->getValue();
                    } else {
                        echo 'Incorrect hero name ' . $nameHero . ' event "Battle of Heroes"'; echo '<br>';
                    }
                } else break;
            }

            $this->sheetTwo->setCellValue('G'.$i, array_sum($power));
            
            $currentPower = $this->sheetTwo->getCell('G'.$i)->getValue();
            $recommendedPower = $this->sheetTwo->getCell('H'.$i)->getValue();

            // Power comparison and background change
            if($currentPower > $recommendedPower)
            {
                $this->sheetTwo->getStyle('G'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => '00ff00')
                        )
                    )
                );
            } else {
                $this->sheetTwo->getStyle('G'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'ff0000')
                        )
                    )
                );
            }
        }

        if($this->sheetTwo->getStyle('A3')->getFill()->getStartColor()->getRGB() == '000000' ) 
        {
            $this->setBackgroundForHeroes(3, 'A', 'F');
        }
        echo 'Инициализация события Битва Героев успешно завершено';echo '<br>';
    } 

    /**
	 *	Search for the power of heroes by matching the name of the hero. 
     *  Calculation of the total power of the team of heroes. 
     *  Comparison of the current and recommended power of the heroes and changing the background of the cell. 
     *  Event "Sky City"
	 *	     
     *  @param array        $heroes
	 */
    public function calculationSkyCity(array $heroes)
    {
        echo 'Инициализация события Небесный Город:';echo '<br>';
        for ($i = 3;  ; $i++) 
        {
            $CellRow = $this->sheetTwo->getCell('L'.$i)->getValue();

            if (empty($CellRow)) {
                break;
            }

            unset($power);
            $power[4];

            //  Current total power
            for($j = 'L'; $j <= 'P'; $j++)
            {
                if (!empty($CellRow)) {
                    $nameHero = $this->sheetTwo->getCell($j.$i)->getValue();
                    $nameHero = trim($nameHero);
                    if (in_array($nameHero, $heroes)) {
                        preg_match('/\d+/', array_shift(array_keys($heroes, $nameHero)), $mathes);
                        $power[] = $this->sheetOne->getCell('C'.array_shift($mathes))->getValue();
                    } else {
                        echo 'Incorrect hero name ' . $nameHero . ' event "Sky City"'; echo '<br>';
                    }
                } else break; 
            }

            $this->sheetTwo->setCellValue('Q'.$i, array_sum($power));
            
            $currentPower = $this->sheetTwo->getCell('Q'.$i)->getValue();
            $recommendedPower = $this->sheetTwo->getCell('R'.$i)->getValue();

            // Power comparison and background change
            if($currentPower > $recommendedPower)
            {
                $this->sheetTwo->getStyle('Q'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => '00ff00')
                        )
                    )
                );
            } else {
                $this->sheetTwo->getStyle('Q'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'ff0000')
                        )
                    )
                );
            }
        }

        if($this->sheetTwo->getStyle('L3')->getFill()->getStartColor()->getRGB() == '000000' ) 
        {
            $this->setBackgroundForHeroes(3, 'L', 'P');
        }
        echo 'Инициализация события Небесный Город успешно завершено';echo '<br>';
    }

    /**
	 *	Search for the power of heroes by matching the name of the hero. 
     *  Calculation of the total power of the team of heroes. 
     *  Comparison of the current and recommended power of the heroes and changing the background of the cell. 
     *  Event "Glory Trials"
	 *	     
     *  @param array        $heroes
	 */
    public function calculationGloryTrials(array $heroes)
    {
        echo 'Инициализация события Испытание Славы:';echo '<br>';
        for ($i = 12;  ; $i++) 
        {
            $CellRow = $this->sheetTwo->getCell('K'.$i)->getValue();

            if (empty($CellRow)) {
                break;
            }

            unset($power);
            $power[4];

            //  Current total power
            for($j = 'K'; $j <= 'O'; $j++)
            {
                if (!empty($CellRow)) {
                    $nameHero = $this->sheetTwo->getCell($j.$i)->getValue();
                    $nameHero = trim($nameHero);
                    if (in_array($nameHero, $heroes)) {
                        preg_match('/\d+/', array_shift(array_keys($heroes, $nameHero)), $mathes);
                        $power[] = $this->sheetOne->getCell('C'.array_shift($mathes))->getValue();
                    } else {
                        echo 'Incorrect hero name ' . $nameHero . ' event "Glory Trials"'; echo '<br>';
                    }
                } else break; 
            }

            $this->sheetTwo->setCellValue('P'.$i, array_sum($power));
            
            $currentPower = $this->sheetTwo->getCell('P'.$i)->getValue();
            $recommendedPower = $this->sheetTwo->getCell('Q'.$i)->getValue();

            // Power comparison and background change
            if($currentPower > $recommendedPower)
            {
                $this->sheetTwo->getStyle('P'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => '00ff00')
                        )
                    )
                );
            } else {
                $this->sheetTwo->getStyle('P'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'ff0000')
                        )
                    )
                );
            }
        }
        echo 'Инициализация события Испытание Славы успешно завершено';echo '<br>';
    }

    /**
	 *	Search for the power of heroes by matching the name of the hero. 
     *  Calculation of the total power of the team of heroes. 
     *  Comparison of the current and recommended power of the heroes and changing the background of the cell. 
     *  Event "Star Trial"
	 *	
     *  @param array        $heroes
	 */
    public function calculationStarTrial(array $heroes)
    {
        echo 'Инициализация события Звездное Испытание:';echo '<br>';
        for ($i = 18;  ; $i++) 
        {
            $CellRow = $this->sheetTwo->getCell('M'.$i)->getValue();

            if (empty($CellRow)) {
                break;
            }

            unset($power);
            $power[4];

            //  Current total power
            for($j = 'M'; $j <= 'Q'; $j++)
            {
                if (!empty($CellRow)) {
                    $nameHero = $this->sheetTwo->getCell($j.$i)->getValue();
                    $nameHero = trim($nameHero);
                    if (in_array($nameHero, $heroes)) {
                        preg_match('/\d+/', array_shift(array_keys($heroes, $nameHero)), $mathes);
                        $power[] = $this->sheetOne->getCell('C'.array_shift($mathes))->getValue();
                    } else {
                        echo 'Incorrect hero name ' . $nameHero . ' event "Star Trial"'; echo '<br>';
                    }
                } else break; 
            }

            $this->sheetTwo->setCellValue('R'.$i, array_sum($power));
            
            $currentPower = $this->sheetTwo->getCell('R'.$i)->getValue();
            $recommendedPower = $this->sheetTwo->getCell('S'.$i)->getValue();

            // Power comparison and background change
            if($currentPower > $recommendedPower)
            {
                $this->sheetTwo->getStyle('R'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => '00ff00')
                        )
                    )
                );
            } else {
                $this->sheetTwo->getStyle('R'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'ff0000')
                        )
                    )
                );
            }
        }
        if($this->sheetTwo->getStyle('M18')->getFill()->getStartColor()->getRGB() == '000000' ) 
        {
            $this->setBackgroundForHeroes(18, 'M', 'Q');
        }
        echo 'Инициализация события Звездное Испытание успешно завершено';echo '<br>';
    }

    /**
	 *	Search for the power of heroes by matching the name of the hero. 
     *  Calculation of the total power of the team of heroes. 
     *  Comparison of the current and recommended power of the heroes and changing the background of the cell. 
     *  Event "Subterra"
	 *	     
     *  @param array        $heroes
	 */
    public function calculationSubterra(array $heroes)
    {
        echo 'Инициализация события Субтерра:';echo '<br>';
        for ($i = 12;  ; $i++) 
        {
            $CellRow = $this->sheetTwo->getCell('B'.$i)->getValue();

            if (empty($CellRow)) {
                break;
            }

            unset($power);
            $power[4];

            //  Current total power
            for($j = 'B'; $j <= 'F'; $j++)
            {
                if (!empty($CellRow)) {
                    $nameHero = $this->sheetTwo->getCell($j.$i)->getValue();
                    $nameHero = trim($nameHero);
                    if (in_array($nameHero, $heroes)) {
                        preg_match('/\d+/', array_shift(array_keys($heroes, $nameHero)), $mathes);
                        $power[] = $this->sheetOne->getCell('C'.array_shift($mathes))->getValue();
                    } else {
                        echo 'Incorrect hero name ' . $nameHero . ' event "Subterra"'; echo '<br>';
                    }
                } else break; 
            }

            $this->sheetTwo->setCellValue('G'.$i, array_sum($power));
            
            $currentPower = $this->sheetTwo->getCell('G'.$i)->getValue();
            $recommendedPower = $this->sheetTwo->getCell('H'.$i)->getValue();

            // Power comparison and background change
            if($currentPower > $recommendedPower)
            {
                $this->sheetTwo->getStyle('G'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => '00ff00')
                        )
                    )
                );
            } else {
                $this->sheetTwo->getStyle('G'.$i)->applyFromArray(
                    array(
                        'fill' => array(
                            'type' => PHPExcel_Style_Fill::FILL_SOLID,
                            'color' => array('rgb' => 'ff0000')
                        )
                    )
                );
            }
        }
        if($this->sheetTwo->getStyle('B12')->getFill()->getStartColor()->getRGB() == '000000' ) 
        {
            $this->setBackgroundForHeroes(12, 'B', 'F');
        }
        echo 'Инициализация события Субтерра успешно завершено';echo '<br>';
    }

    /**
	 *	Linear RGB interpolation based on percentage. 
     *  Changing the background of a cell (similar to the conditional formatting rule with a three-color scale in Excel)
	 *	     
	 */
    public function setBackgroundRGB()
    {
        $namesOfColumnsRGB = ['K', 'L', 'M', 'N', 'Q', 'R'];
        for ( $i = 3; ; $i++) {
            $CellRow = $this->sheetOne->getCell('A'.$i)->getValue();
            if (!empty($CellRow) && $CellRow != 'Итог:') {
                for ($j = 0; $j <= count($namesOfColumnsRGB)-1; $j++) {
                    $value = $this->sheetOne->getCell($namesOfColumnsRGB[$j].$i)->getValue()*100;
                    $b = 0;
                    switch ($value) {
                        case 0:
                            $r = 255;
                            $g = 0;
                            break;
                        case ($value > 0 && $value < 50):
                            $r = 255;
                            $g = 255*$value/50;
                            break;
                        case 50:
                            $r = 255;
                            $g = 255;
                            break;
                        case ($value > 50 && $value < 100):
                            $r = 255-(255-0)*($value-50)/(100-50);
                            $g = 255;
                            break;
                        case 100:
                            $r = 0;
                            $g = 255;
                            break;
                        default:
                            echo 'Входящее число не принадлежит диапазону 0-100 %';
                            break;
                    }

                    $color = sprintf("%02x%02x%02x", $r, $g, $b);

                    $this->sheetOne->getStyle($namesOfColumnsRGB[$j].$i)->applyFromArray(
                        array(
                            'fill' => array(
                                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                'color' => array('rgb' => $color)
                            )
                        )
                    );
                }
            } else break;
        }
    }

    /**
	 *	Sets a random background for heroes used more than once.
	 *	    
     *  @param int          $i                  Number is used as the first expression of the first loop
     *  @param string       $startCell          The letter is used as the first expression of the second loop
     *  @param string       $endCell            Еhe letter is used as the second expression of the second loop
	 */
    public function setBackgroundForHeroes(int $i, string $startCell, string $endCell)
    {
        for ($i = $i;  ; $i++) 
        {
            $CellRow = $this->sheetTwo->getCell($startCell.$i)->getValue();

            if (empty($CellRow)) {
                break;
            }

            for($j = $startCell; $j <= $endCell; $j++)
            {
                if (!empty($CellRow)) {
                    $nameHeroes[$j.$i] = $this->sheetTwo->getCell($j.$i)->getValue();
                } 
            }
        }

        $countHeroes = array_filter(array_count_values($nameHeroes), function($var) {
            return $var > 1;
            }
        );
        
        $arrayColors = array();
        foreach ($countHeroes as $key => $value) 
        {
            $keys = array_keys($nameHeroes, $key);

            color:
            $color = sprintf("%02x%02x%02x", rand(0,255), rand(0,255), rand(0,255)); 
        
            if(in_array($color, $arrayColors) == false)
            {
                $arrayColors[] = $color;
                
                foreach ($keys as $value) {
                    $this->sheetTwo->getStyle($value)->applyFromArray(
                        array(
                            'fill' => array(
                                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                'color' => array('rgb' => $color)
                            )
                        )
                    );
                }
            } else 
            {
                goto color;
            }
        }
    }
}

    $magicRush = new MagicRush;
    $heroes = $magicRush->getHeroes();
    $magicRush->calculationBattleOfHeroes($heroes);
    $magicRush->calculationSkyCity($heroes);
    $magicRush->calculationGloryTrials($heroes);
    $magicRush->calculationStarTrial($heroes);
    $magicRush->calculationSubterra($heroes);
    $magicRush->setBackgroundRGB();

