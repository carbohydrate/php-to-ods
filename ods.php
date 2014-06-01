<?php
class ods {
    //files
    public $content;
    //private $meta;
    //private $settings;
    //private $style;
    //private $path;
    public $pages;
    public $page;
    public $rows;

    function __construct() {
        $this->content = array();
        //$this->meta = array();
        //$this->settings = array();
        //$this->style = array();
        //$this->path = 'uploads/file.ods';
        $this->page = array();
        $this->rows = array();
    }
    public function addCell($page, $row, $column, $value, $type, $width = 1) {
        if ($type == 'float') {
            $this->pages[$page]['rows'][$row][$column]['cellValue'] = array('office:value-type'=>$type,'office-value'=>$value, 'text:p'=>$value, 'xml:id'=>$column);
        } else {
            $this->pages[$page]['rows'][$row][$column]['cellValue'] = array('office:value-type'=>$type, 'text:p'=>$value, 'xml:id'=>$column);
        }
    }


    public function getContent() {
        $this->content[0] = '<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2"><office:scripts/>';
        $this->content[1] = '<office:automatic-styles>';
        $this->content[2] = '<style:style style:name="co1" style:family="table-column">';
        $this->content[3] = '<style:table-column-properties fo:break-before="auto" style:column-width="0.889in"/>';
        $this->content[4] = '</style:style>';
        $this->content[5] = '<style:style style:name="ro1" style:family="table-row">';
        $this->content[6] = '<style:table-row-properties style:row-height="0.178in" fo:break-before="auto" style:use-optimal-row-height="true"/>';
        $this->content[7] = '</style:style>';
        $this->content[8] = '<style:style style:name="ta1" style:family="table" style:master-page-name="Default">';
        $this->content[9] = '<style:table-properties table:display="true" style:writing-mode="lr-tb"/>';
        $this->content[10] = '</style:style>';
        $this->content[11] = '</office:automatic-styles>';

        $this->content[12] = '<office:body>';
        $this->content[13] = '<office:spreadsheet>';
        $this->content[14] = '<table:table table:name="Sheet1" table:style-name="ta1">';
        $this->content[15] = '';


        $this->content[16] = '</table:table>';
        $this->content[17] = '</office:spreadsheet>';
        $this->content[18] = '</office:body>';
        $this->content[19] = '</office:document-content>';
        $this->prepare();
        //print_r($this->content);
        $string = '';
        foreach ($this->content as $value) {
            $string .= $value;
        }
        return $string;
    }

    public function prepare() {
        $pages = $this->pages;
        $string = '';
        foreach ($pages as $rowz) {
            foreach ($rowz as $colz) {
                foreach ($colz as $key => $a) {
                    ksort($a);  //make sure columns are in correct order. incase user inserted column 4 before 2.
                    $string .= '<table:table-row table:style-name="ro1" xml:id="' . $key . '">';
                    foreach ($a as $value) {
                        //print_r($value);
                        $string .= '<table:table-cell office:value-type="' . $value['cellValue']['office:value-type'] . '"><text:p>' . $value['cellValue']['text:p'] . '</text:p></table:table-cell>';
                    }
                    //$string .= '</table:table-cell>';
                }
                $string .= '</table:table-row>';
            }
            $this->content[15] .= $string;
        }

    }

    public function exportOds($path) {
        $zip = new ZipArchive();
        if ($zip->open($path, ZipArchive::CREATE)!==TRUE) {
            exit('Can\'t Open File');
        }
        $zip->addFromString('content.xml', $this->getContent());
        $zip->addFromString('', $this->getMeta());
        //echo "numfiles: " . $zip->numFiles . "\n";
        //echo "status:" . $zip->status . "\n";
        $zip->close();

    }
}
header('Content-Type: text/plain');
$obj = new ods;

//$obj->exportOds($obj, $path);
//$obj = $obj->getContent();
$obj->addCell(1,1,1,'test','string');
$obj->addCell(1,1,2,'f','string');
$obj->addCell(1,1,3,'row2','string');
$obj->addCell(1,1,4,'row2spot1','string');
$obj->addCell(1,1,5,'row2','string');
//$obj->addCell(2,1,2,'fuck','string');
//$obj->prepare();

//print_r($obj->content);

$path = 'uploads/file.ods';
$obj->exportOds($path);


?>
