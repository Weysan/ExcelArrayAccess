<?php
 
/**
 * Va permettre de gérer un fichier excel comme un teableau PHP
 * On pourra utiliser foreach() - ArrayIterator
 * On pourra utiliser la nomenclature des tableaux $arr[] - ArrayAccess
 * Les index doivent être utilisés uniquement en entier (0 à XX)
 *
 * @author Raphael GONCALVES <contact@raphael-goncalves.fr>
 */
class ExcelArray extends ArrayIterator implements ArrayAccess
{
    private $container = array();
     
    private $phpExcel;
     
    private $filename;
     
    private $modification = false;
     
    public function __construct($filename = null, $sheet = 0)
    {
        $this->filename = $filename;
         
        if(file_exists($filename)){
            /* on charge le contenu du fichier dans l'objet phpExcel */
            $inputFileType = PHPExcel_IOFactory::identify($filename);
            $objReader = PHPExcel_IOFactory::createReader($inputFileType);
            $this->phpExcel = $objReader->load($filename);
             
            /* on implémente dans le tableau */
            //  Get worksheet dimensions
            $sheet = $this->phpExcel->getSheet($sheet); 
            $highestRow = $sheet->getHighestRow(); 
            $highestColumn = $sheet->getHighestColumn();
 
            //  Loop through each row of the worksheet in turn
            for ($row = 1; $row <= $highestRow; $row++){ 
                //  Read a row of data into an array
                $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row,
                                                NULL,
                                                TRUE,
                                                FALSE);
                $this->container[] = $rowData[0];
            }
        } else {
            $this->phpExcel = new PHPExcel();
            $this->phpExcel->setActiveSheetIndex($sheet);
        }
         
        parent::__construct($this->container);
    }
     
    /**
     * Va assigner une valeur à un index.
     * @param mixed $offset
     * @param mixed $value
     */
    public function offsetSet($offset, $value) {
        if (is_null($offset)) {
            $this->container[] = $value;
        } else {
            $this->container[$offset] = $value;
        }
         
        /* on assigne les valeurs à l'objet excel aussi */
        if(is_array($value)){
            foreach($value as $key => $data){
                $index = $this->getLetter($key).($offset+1); //il n'y a pas de 0 en index dans les lignes
                $this->phpExcel->getActiveSheet()->setCellValue($index, $data);
            }
        } else {
            $this->phpExcel->getActiveSheet()->setCellValue('A'.$offset, $value);
        }
         
        $this->modification = true;
    }
 
    /**
     * Vérifier l'existance d'une donnée.
     * @param mixed $offset
     * @return boolean
     */
    public function offsetExists($offset) {
        return isset($this->container[$offset]);
    }
 
    /**
     * Permet de supprimer des données dans le fichier.
     * @param mixed $offset
     */
    public function offsetUnset($offset) {
        unset($this->container[$offset]);
         
        //on supprime la ligne qu'on supprime
        $lastRow = $this->phpExcel->getActiveSheet()->getHighestRow();
        for ($row = 1; $row <= $lastRow; $row++) {
            $cell = $this->phpExcel->getActiveSheet()->setCellValue($offset.$row, null);
        }
         
        $this->modification = true;
         
    }
 
    /**
     * Retourne la valeur du tableau.
     * @param mixed $offset
     * @return mixed
     */
    public function offsetGet($offset) {
        return isset($this->container[$offset]) ? $this->container[$offset] : null;
    }
     
    /**
     * Permet de retrouver une lettre en fonction de sa position.
     * @param integer $number
     * @return string
     */
    private function getLetter($number){
        //vérifier les données en entrée de fonction.
        assert(is_int($number), '$number doit être un nombre entier.');
         
        $letter = 'A';
        while($number != ord(strtoupper($letter)) - ord('A')){
            $letter++;
        }
         
        return $letter;
    }
     
    /**
     * Permet de retrouver un index de tableau en fonction de la lettre.
     * @param string $letter
     * @return integer position de la lettre
     */
    private function getNumber($letter){
        //vérifier les données en entrée de fonction.
        assert(is_string($letter), '$number doit être une chaine de caractère.');
        return ord(strtoupper($letter)) - ord('A');
    }
     
    /**
     * Si une modification des données à été réalisée, on sauvegarde le fichier.
     */
    public function __destruct()
    {
        if($this->modification && !is_null($this->filename)){
            $objWriter = PHPExcel_IOFactory::createWriter($this->phpExcel, 'Excel2007');
            $objWriter->save($this->filename);
        }
    }
}
