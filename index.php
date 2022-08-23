<?php




require 'vendor/autoload.php';

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("./question_template.xlsx");

$worksheet = $spreadsheet->getActiveSheet();
$worksheetArray = $worksheet->toArray();
echo '<table style="width:100%"  border="1">';
echo '<tr align="center">';
echo '<td>Question</td>';
echo '<td>Option 1</td>';
echo '<td>Option 2</td>';
echo '<td>Option 3</td>';
echo '<td>Option 4</td>';
echo '<td>Answer</td>';
echo '</tr>';

$store_img_arr = array();
$cell_index = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N');

for($i=0; $i<count($worksheet->getDrawingCollection()); $i++)
{
    $coordinates = $worksheet->getDrawingCollection()[$i]->getCoordinates();

    $drawing = $worksheet->getDrawingCollection()[$i];
    $zipReader = fopen($drawing->getPath(), 'r');
    $imageContents = '';
    $extension = $drawing->getExtension();
    while (!feof($zipReader)) {
        $imageContents .= fread($zipReader, 1024); 
    }
  $filename = 'images/'.$i.time().'.'.$extension;
    file_put_contents($filename, $imageContents);
    fclose($zipReader);
    $store_img_arr[$coordinates] = $filename;
}
//print_r($store_img_arr);
$k = 0;
for($i = 1; $i < count($worksheetArray); $i++) {
    for($j = 1; $j <= 10; $j++) {

            if(isset( $store_img_arr[$cell_index[$j].($i+1)])){
                $worksheetArray[$i][$j] =  $worksheetArray[$i][$j]."#@ImGq@#".$store_img_arr[$cell_index[$j].($i+1)];
            }
            else
            {
                $worksheetArray[$i][$j] =  $worksheetArray[$i][$j];

            }
 
    }
}





for($i = 1; $i < count($worksheetArray); $i++) {
$j = 1;
     echo '<tr align="center">';


     for($c=1; $c<=5; $c++)
     {


        $question = explode("#@ImGq@#", $worksheetArray[$i][$j]);


        
        if(count($question) == 1)
        {
            
                echo '<td>' . $question[0] . '</td>';
             
        }
        elseif(count($question) == 2)
        {
            $question_text = $question[0];

            $question_img = '<img  height="150px" width="150px"   src="'.$question[1].'"/>';

            if(strlen($question_text) == 0)
            {
                $print_question = $question_img;
            }
            else
            {
                $print_question = str_replace("<<q@img>>", $question_img ,$question_text);
            }
            echo '<td>' . $print_question . '</td>';

        }
       


  
        $j++;  
    } 

    



     echo '<td>' . $worksheetArray[$i][9] . '</td>';
     $j++;

     echo '</tr>';
 
}