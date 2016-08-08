<?php
require __DIR__ . '/vendor/autoload.php';
error_reporting(E_ERROR | E_PARSE);



$phpWord = new \PhpOffice\PhpWord\PhpWord();

$TAITAN = 'https://taitan.datasektionen.se';
$ME_URL = $_ENV['APP_URL']."?path=";

$json = file_get_contents($TAITAN.$_GET['path']);
$obj = json_decode($json);



$file = $obj->title.'.docx';





header("Content-Description: File Transfer");
header('Content-Disposition: inline; filename="' . $file . '"');
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
header('Expires: 0');
#$xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
#$xmlWriter->save("php://output");
\PhpOffice\PhpWord\Settings::setOutputEscapingEnabled(true);
// New Word document

// New section
$section = $phpWord->addSection();
// Define styles
$phpWord->addTitleStyle(1, array('size' => 35, 'bold' => true, 'color' => '#e83d84'));

// Adding Text element with font customized using named font style...
$pStyle = 'p';
$phpWord->addFontStyle(
    $pStyle,
    array('name' => 'Arial', 'size' => 16, 'color' => '757575', 'bold' => false)
);

// Arc
$section->addTitle($obj->title, 1);
$section->addText(htmlspecialchars_decode(strip_tags($obj->body)), $pStyle);

#\PhpOffice\PhpWord\Shared\Html::addHtml($section, htmlspecialchars_decode($obj->body)) ;
if ($obj->sidebar) {
  $sidebar = $phpWord->addSection();
  $sidebar->addTitle("Sidebar", 1);
  $sidebar->addText(strip_tags($obj->sidebar), $pStyle);
}

$nav = $phpWord->addSection();
$nav->addTitle("Navigation", 1);
foreach ($obj->nav as $anchor) {
  $nav->addLink($ME_URL.$anchor->slug, $anchor->title);
  if ($anchor->nav) {
    foreach ($anchor->nav as $subnav) {
      $nav->addLink($ME_URL.$subnav->slug, "  ".$subnav->title);
    }
  }
}


$xmlWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$xmlWriter->save("php://output");
/*
// New section
$section = $phpWord->addSection();
// Define styles
$phpWord->addTitleStyle(1, array('size' => 14, 'bold' => true));
// Arc
$section->addTitle('Arc', 1);
$section->addShape(
    'arc',
    array(
        'points'  => '-90 20',
        'frame'   => array('width' => 120, 'height' => 120),
        'outline' => array('color' => '#333333', 'weight' => 2, 'startArrow' => 'oval', 'endArrow' => 'open'),
    )
);
// Curve
$section->addTitle('Curve', 1);
$section->addShape(
    'curve',
    array(
        'points'    => '1,100 200,1 1,50 200,50',
        'connector' => 'elbow',
        'outline'   => array(
            'color'      => '#66cc00',
            'weight'     => 2,
            'dash'       => 'dash',
            'startArrow' => 'diamond',
            'endArrow'   => 'block',
        ),
    )
);
// Line
$section->addTitle('Line', 1);
$section->addShape(
    'line',
    array(
        'points'  => '1,1 150,30',
        'outline' => array(
            'color'      => '#cc00ff',
            'line'       => 'thickThin',
            'weight'     => 3,
            'startArrow' => 'oval',
            'endArrow'   => 'classic',
        ),
    )
);
// Polyline
$section->addTitle('Polyline', 1);
$section->addShape(
    'polyline',
    array(
        'points'  => '1,30 20,10 55,20 75,10 100,40 115,50, 120,15 200,50',
        'outline' => array('color' => '#cc6666', 'weight' => 2, 'startArrow' => 'none', 'endArrow' => 'classic'),
    )
);
// Rectangle
$section->addTitle('Rectangle', 1);
$section->addShape(
    'rect',
    array(
        'roundness' => 0.2,
        'frame'     => array('width' => 100, 'height' => 100, 'left' => 1, 'top' => 1),
        'fill'      => array('color' => '#FFCC33'),
        'outline'   => array('color' => '#990000', 'weight' => 1),
        'shadow'    => array(),
    )
);
// Oval
$section->addTitle('Oval', 1);
$section->addShape(
    'oval',
    array(
        'frame'     => array('width' => 100, 'height' => 70, 'left' => 1, 'top' => 1),
        'fill'      => array('color' => '#33CC99'),
        'outline'   => array('color' => '#333333', 'weight' => 2),
        'extrusion' => array(),
    )
);*/

?>
