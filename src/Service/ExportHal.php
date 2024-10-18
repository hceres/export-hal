<?php

namespace App\Service;

use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\HttpFoundation\JsonResponse;

class ExportHal
{

    private const WIDTH_ADAPTATOR = 0.7;

    /**
     * @param string $idcoll
     * @param string $dateDeb
     * @param string $dateFin
     * @param string $soulignauteur
     * @param string $equipelabo
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public static function getResult(string $recherche, string $idshal, string $idcoll, string $dateDeb, string $dateFin, string $soulignauteur, string $equipelabo): string
    {

        ini_set('memory_limit', '1024M');
        ini_set('max_execution_time', '240');
        date_default_timezone_set('Europe/London');
        define('EOL', (PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

        $debug = false;
        if ($debug)
            $h = fopen("debug.log", "w");

        $noticesTraitees = array();  // stockage de toutes les publis traitées, pour remplir la section des publis non ventilées
        $nbNoticeTraite = 0;
        $liste1 = "halId_s,title_s,authFullName_s,authLastName_s,authFirstName_s,producedDate_s,producedDateY_i,docType_s,inra_publicVise_local_s,peerReviewing_s,invitedCommunication_s,subTitle_s,bookTitle_s,journalTitle_s,volume_s,issue_s,page_s,publisher_s,doiId_s,uri_s,arxivId_s,biorxivId_s,authorityInstitution_s,number_s,serie_s,conferenceTitle_s,city_s,country_s,conferenceStartDate_s,conferenceEndDate_s,lectureName_s,reportType_s,lectureType_s,submitType_s,openAccess_bool,wosId_s,pubmedId_s,audience_s,otherType_s,authQuality_s,authIdFullName_fs";
        $liste2 = "inra_otherType_Other_local_s,inra_otherType_Art_local_s,inra_otherType_Undef_local_s,inra_otherType_Douv_local_s,inra_otherType_Comm_local_s,inra_reportType_local_s,popularLevel_s,authIdHasPrimaryStructure_fs,linkExtId_s,language_s,description_s,abstract_s,keyword_s";
        $doctypeQueryHal = "ART+OR+COMM+OR+POSTER+OR+OUV+OR+COUV+OR+PROCEEDINGS+OR+BLOG+OR+ISSUE+OR+NOTICE+OR+TRAD+OR+PATENT+OR+OTHER+OR+UNDEFINED+OR+REPORT+OR+THESE+OR+HDR+OR+LECTURE+OR+VIDEO+OR+SON+OR+IMG+OR+MAP+OR+SOFTWARE";
        $listeChamps = $liste1 . ',' . $liste2;
        $myXls = "HAL-production";
        // Récuperation des données du formulaire
        $collection = htmlspecialchars(strip_tags($idcoll), ENT_QUOTES);
        $anneedeb = htmlspecialchars(strip_tags($dateDeb), ENT_QUOTES);
        $anneefin = htmlspecialchars(strip_tags($dateFin), ENT_QUOTES);
        $soulignAuteur = 1;  // 1 pour cocher, 0 sinon
        // id AuréHAL, le nom ou l'acronyme de votre unité, selon que vous souhaitez mettre en évidence le nom des auteurs de l'unité. Exemple 928~ECOBIO~575446
        $equipeLabo = strtoupper(htmlspecialchars(strip_tags($equipelabo), ENT_QUOTES));
        $idshal = htmlspecialchars(strip_tags($idshal), ENT_QUOTES);
        $recherche = htmlspecialchars(strip_tags($recherche), ENT_QUOTES);

        $notices = [];

        if (isset($collection)) {
            $collection = strtoupper($collection);
        }

        if ($recherche == "code" && $equipeLabo == "" && $collection != "") {
            $equipeLabo = $collection;
            $myXls .='-'.$collection;
        } elseif ($recherche == "code" && $collection == "" && $equipeLabo != "") {
            $myXls .='-'.$equipeLabo;
        } elseif ($recherche == "code" && $collection != "" && $equipeLabo != "") {
            $myXls .='-'.$equipeLabo."-".$collection;
        } else {
            $myXls .="-".$idshal;
        }

        $myXls.=".xlsm";

        include_once("../src/Helper/exporthal/functions.php");
        include_once("../src/Helper/exporthal/functions_hceres.php");


        if (isset($recherche) && $recherche == "code") {

            if (isset($collection) && !empty($collection)) {

                $url = "https://api.archives-ouvertes.fr/search/?wt=json&q=collCode_s:" . $collection . "&rows=100000&fq=producedDateY_i:[" . $anneedeb . "+TO+" . $anneefin . "]&sort=producedDateY_i%20desc&fl=" . $listeChamps;
                $urlHAL = "https://hal.archives-ouvertes.fr/search/index/?q=collCode_s:" . $collection . "+producedDateY_i:[" . $anneedeb . "+TO+" . $anneefin . "]&sort=producedDateY_i%20desc&submit=&docType_s=".$doctypeQueryHal."&submitType_s=notice+OR+file+OR+annex&rows=300";
            } else {
                $url = "https://api.archives-ouvertes.fr/search/?wt=json&q=labStructAcronym_s:" . $equipeLabo . "&rows=100000&fq=producedDateY_i:[" . $anneedeb . "+TO+" . $anneefin . "]&sort=producedDateY_i%20desc&fl=" . $listeChamps;
                $urlHAL = "https://hal.archives-ouvertes.fr/search/index/?q=labStructAcronym_s:" . $equipeLabo . "+producedDateY_i:[" . $anneedeb . "+TO+" . $anneefin . "]&sort=producedDateY_i%20desc&submit=&docType_s=".$doctypeQueryHal."&submitType_s=notice+OR+file+OR+annex&rows=300";
            }
            if ($anneedeb == "") $annedeb = "*";
            if ($anneefin == "") $anneefin = "*";


        } elseif (isset($recherche) && $recherche == "aurehal" && isset($idshal) && !empty($idshal)) {
            $ids = explode("-",$idshal);
            $structureQuerySearch = "";
            foreach ($ids as $index => $id) {
                $structureQuerySearch .= $index > 0 ? "%20||%20" : "";
                $structureQuerySearch .= $id;
            }
            $url = "http://api.archives-ouvertes.fr/search/?wt=json&q=structId_i:(" . $structureQuerySearch . ")&rows=100000&fq=producedDateY_i:[" . $anneedeb . "+TO+" . $anneefin . "]&sort=producedDateY_i%20desc&fl=" . $listeChamps;
            $urlHAL = "https://hal.archives-ouvertes.fr/search/index/?q=structId_i:(" . $structureQuerySearch . ")+producedDateY_i:[" . $anneedeb . "+TO+" . $anneefin . "]&sort=producedDateY_i%20desc&submit=&docType_s=ART+OR+COMM+OR+POSTER+OR+OUV+OR+COUV+OR+DOUV+OR+PATENT+OR+OTHER+OR+UNDEFINED+OR+REPORT+OR+CREPORT+OR+THESE+OR+HDR+OR+LECTURE+OR+MEM+OR+VIDEO+OR+SON+OR+IMG+OR+MAP+OR+SOFTWARE&submitType_s=notice+OR+file+OR+annex&rows=300";
        } else {
            return json_encode("error");
        }

        //$linkInrae = '<a href="' . $urlHAL . '" target="_blank">Lancer la recherche sur HAL - INRAE</a>';

        $ch = curl_init();
        curl_setopt($ch, CURLOPT_URL, $url);
        curl_setopt($ch, CURLOPT_HEADER, 0);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
        if (isset ($_SERVER["HTTPS"]) && $_SERVER["HTTPS"] == "on") {
            curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, TRUE);
            curl_setopt($ch, CURLOPT_CAINFO, "cacert.pem");
        }
        $resultats = curl_exec($ch);
        curl_close($ch);

        $data = json_decode($resultats, true);
        $notices = $data["response"]["docs"];

        if (count($notices) == 0) {
            $retourne[1] = "Aucune notice ne correspond à ces critères. Veuillez saisir un code de collection correct SVP.";
            $retourne[2] = "Erreur";
            //$error = true;
            echo json_encode($retourne);
            exit;
        } else {
            //$msg = count($notices) . " notices correspondent aux critères";
            //$error = false;
        }

        $equipes = "";
        if (isset($equipeLabo) && !empty($equipeLabo)) {
            $equipes = getAffiliations($equipeLabo);  // on récupère toutes les valeurs à comparer avec les auteurs des notices, liste de docid de structure 928~1050~...
        }

        $oSpreadsheet = new Spreadsheet();

        // Setting the default parameters
        $oSpreadsheet->getDefaultStyle()->getFont()->setName('Century Gothic');
        $oSpreadsheet->getDefaultStyle()->getFont()->setSize(11);
        $oSpreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(13.8);

        $coverSheet = $oSpreadsheet->getActiveSheet();

        $coverSheet->setTitle("Production HAL");

        $coverSheet
            ->getStyle('A1:AZ200')
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('ffffff');

        $coverSheet->getRowDimension(1)->setRowHeight(13.5);
        $coverSheet->getRowDimension(2)->setRowHeight(13.5);
        $coverSheet->getRowDimension(3)->setRowHeight(24);
        $coverSheet->getRowDimension(4)->setRowHeight(23.4);
        $coverSheet->getRowDimension(5)->setRowHeight(17.4);
        $coverSheet->getRowDimension(6)->setRowHeight(13.5);
        $coverSheet->getRowDimension(7)->setRowHeight(18.6);
        $coverSheet->getRowDimension(8)->setRowHeight(18.6);
        $coverSheet->getRowDimension(9)->setRowHeight(18.6);
        $coverSheet->getRowDimension(10)->setRowHeight(18.6);
        $coverSheet->getRowDimension(11)->setRowHeight(18.6);
        $coverSheet->getRowDimension(12)->setRowHeight(18.6);

        $coverSheet->getDefaultColumnDimension()->setWidth(10.67+self::WIDTH_ADAPTATOR);

        $coverSheet->getColumnDimension('A')->setWidth(6.11+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('B')->setWidth(51.89+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('D')->setWidth(2.22+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('E')->setWidth(38.11+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('G')->setWidth(2.22+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('H')->setWidth(23.11+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('J')->setWidth(2.22+self::WIDTH_ADAPTATOR);
        $coverSheet->getColumnDimension('K')->setWidth(24.89+self::WIDTH_ADAPTATOR);

        $coverSheet->setCellValue('C4', 'Département d\'évaluation de la recherche');
        $coverSheet->getStyle('C4')->getFont()->getColor()->setARGB('ff7030a0');
        $coverSheet->getStyle('C4')->getFont()->setSize(18);
        $coverSheet->getStyle('C4')->getFont()->setBold(true);

        $coverSheet->setCellValue('C5','Vague D : campagne d\'évaluation 2023-2024');
        $coverSheet->getStyle('C5')->getFont()->getColor()->setARGB('ff7030a0');
        $coverSheet->getStyle('C5')->getFont()->setSize(14);
        $coverSheet->getStyle('C5')->getFont()->setBold(true);

        $coverSheet->setCellValue('C8','Laboratoire : '.strtoupper($equipeLabo));
        $coverSheet->getStyle('C8')->getFont()->getColor()->setARGB('ff7030a0');
        $coverSheet->getStyle('C8')->getFont()->setSize(14);
        $coverSheet->getStyle('C8')->getFont()->setBold(true);

        $coverSheet->setCellValue('C9','Code collection : '.$collection);
        $coverSheet->getStyle('C9')->getFont()->getColor()->setARGB('ff7030a0');
        $coverSheet->getStyle('C9')->getFont()->setSize(14);
        $coverSheet->getStyle('C9')->getFont()->setBold(true);

        $coverSheet->setCellValue('C10','Identifiant AuréHAL : '.$idshal);
        $coverSheet->getStyle('C10')->getFont()->getColor()->setARGB('ff7030a0');
        $coverSheet->getStyle('C10')->getFont()->setSize(14);
        $coverSheet->getStyle('C10')->getFont()->setBold(true);

        $coverSheet->setCellValue('B13','Publications');
        $coverSheet->getStyle('B13')->getFill()->getStartColor()->setARGB('ffb1a0c7');
        $coverSheet->getStyle('B13')->getFont()->setBold(true);

        $coverSheet->setCellValue('E13','Documents non publiés');
        $coverSheet->getStyle('E13')->getFill()->getStartColor()->setARGB('ff92cddc');
        $coverSheet->getStyle('E13')->getFont()->setBold(true);

        $coverSheet->setCellValue('H13','Travaux universitaires');
        $coverSheet->getStyle('H13')->getFill()->getStartColor()->setARGB('fffabf8f');
        $coverSheet->getStyle('H13')->getFont()->setBold(true);

        $coverSheet->setCellValue('K13','Données de recherche');
        $coverSheet->getStyle('K13')->getFill()->getStartColor()->setARGB('ffda9694');
        $coverSheet->getStyle('K13')->getFont()->setBold(true);

        $coverSheet->getStyle('A1:A50')->getAlignment()->setHorizontal('right');

        $coverSheet->setCellValue('B14','   Article dans une revue');
        $coverSheet->getCell('B14')->getHyperlink()->setUrl("sheet://'Article dans une revue'!A3");

        $coverSheet->setCellValue('B15','   Communication dans un congres');
        $coverSheet->getCell('B15')->getHyperlink()->setUrl("sheet://'Communication dans un congres'!A3");

        $coverSheet->setCellValue('B16','   Poster');
        $coverSheet->getCell('B16')->getHyperlink()->setUrl("sheet://'Poster'!A3");

        $coverSheet->setCellValue('B17','   Proceedings/Recueil des communications');
        $coverSheet->getCell('B17')->getHyperlink()->setUrl("sheet://'Proceedings Recueil des comm.'!A3");

        $coverSheet->setCellValue('B18','   No spécial de revue/special issue');
        $coverSheet->getCell('B18')->getHyperlink()->setUrl("sheet://'No special de revue'!A3");

        $coverSheet->setCellValue('B19','   Ouvrage (y compris édition critique et traduction)');
        $coverSheet->getCell('B19')->getHyperlink()->setUrl("sheet://'Ouvrage'!A3");

        $coverSheet->setCellValue('B20','   Chapitre ouvrage');
        $coverSheet->getCell('B20')->getHyperlink()->setUrl("sheet://'Chapitre ouvrage'!A3");

        $coverSheet->setCellValue('B21','   Article de blog scientifique');
        $coverSheet->getCell('B21')->getHyperlink()->setUrl("sheet://'Article de blog scientifique'!A3");

        $coverSheet->setCellValue('B22','   Notice d\'encyclopedie ou dictionnaire');
        $coverSheet->getCell('B22')->getHyperlink()->setUrl("sheet://'Not. encyclopedie dictionnaire'!A3");

        $coverSheet->setCellValue('B23','   Traduction');
        $coverSheet->getCell('B23')->getHyperlink()->setUrl("sheet://'Traduction'!A3");

        $coverSheet->setCellValue('B24','   Brevet');
        $coverSheet->getCell('B24')->getHyperlink()->setUrl("sheet://'Brevet'!A3");

        $coverSheet->setCellValue('B25','   Autre publication');
        $coverSheet->getCell('B25')->getHyperlink()->setUrl("sheet://'Autre publication'!A3");

        $coverSheet->setCellValue('E14','   Pré-publication, Document de travail');
        $coverSheet->getCell('E14')->getHyperlink()->setUrl("sheet://'Preprint, Working Paper'!A3");

        $coverSheet->setCellValue('E15','   Rapport');
        $coverSheet->getCell('E15')->getHyperlink()->setUrl("sheet://'Rapport'!A3");

        $coverSheet->setCellValue('H14','   These');
        $coverSheet->getCell('H14')->getHyperlink()->setUrl("sheet://'These'!A3");

        $coverSheet->setCellValue('H15','   HDR');
        $coverSheet->getCell('H15')->getHyperlink()->setUrl("sheet://'HDR'!A3");

        $coverSheet->setCellValue('H16','   Cours');
        $coverSheet->getCell('H16')->getHyperlink()->setUrl("sheet://'Cours'!A3");

        $coverSheet->setCellValue('K14','   Media');
        $coverSheet->getCell('K14')->getHyperlink()->setUrl("sheet://'Media'!A3");

        $coverSheet->setCellValue('K15','   Logiciel');
        $coverSheet->getCell('K15')->getHyperlink()->setUrl("sheet://'Logiciel'!A3");

        $drawing = new Drawing();
        $drawing->setName('Logo');
        $drawing->setDescription('Logo');
        $drawing->setPath('../public/assets/img/logo_hceres.jpg');
        $drawing->setHeight(185);
        $drawing->setCoordinates('B2');
        $drawing->setWorksheet($coverSheet);


        $macroSheet = $oSpreadsheet->createSheet();

        $macroSheet->setTitle("Reperer equipes et doctorants");

        $macroSheet->setCellValue('B2','Cet onglet vous propose d\'automatiser le renseignement des équipes et des doctorants dans les colonnes prévues dans chaque onglet contenant des' ."\n".
'productions.' ."\n\n".

'Pour activer le renseignement des colonnes, suivez ces 3 étapes :');
        $macroSheet->getStyle('B2')->getFont()->getColor()->setARGB('ff7030a0');
        $macroSheet->getStyle('B2')->getFont()->setSize(12);
        $macroSheet->getStyle('B2')->getFont()->setBold(true);
        $macroSheet->getStyle('B2')->getAlignment()->setWrapText(true);
        $macroSheet->getStyle('B2')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('B2')->getFill()->getStartColor()->setARGB('ffb1a0c7');
        $macroSheet->getStyle('B2')->getFont()->setBold(true);

        $macroSheet->mergeCells('B2:H3');
        $macroSheet->mergeCells('B5:H5');
        $macroSheet->mergeCells('B6:H6');
        $macroSheet->mergeCells('B12:D12');
        $macroSheet->mergeCells('F12:H12');

//      $macroSheet->removeRow(4);

        $macroSheet->getDefaultColumnDimension()->setWidth(10.1+self::WIDTH_ADAPTATOR);

        $macroSheet->getColumnDimension('A')->setWidth(4+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('B')->setWidth(20.6+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('C')->setWidth(18.3+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('D')->setWidth(25.1+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('E')->setWidth(10.1+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('F')->setWidth(26.3+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('G')->setWidth(22.1+self::WIDTH_ADAPTATOR);
        $macroSheet->getColumnDimension('H')->setWidth(25.1+self::WIDTH_ADAPTATOR);

        $macroSheet->getRowDimension(2)->setRowHeight(40.2);
        $macroSheet->getRowDimension(3)->setRowHeight(40.2);
        $macroSheet->getRowDimension(4)->setRowHeight(0);
        $macroSheet->getRowDimension(5)->setRowHeight(36.8);
        $macroSheet->getRowDimension(6)->setRowHeight(35.3);
        $macroSheet->getRowDimension(7)->setRowHeight(33);
        $macroSheet->getRowDimension(9)->setRowHeight(15.5);
        $macroSheet->getRowDimension(10)->setRowHeight(73.2);
        $macroSheet->getRowDimension(12)->setRowHeight(32);
        $macroSheet->getRowDimension(13)->setRowHeight(27);

        $macroSheet
            ->getStyle('A1:H11')
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('ffffff');

        $macroSheet
            ->getStyle('A1:A200')
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('ffffff');

        $macroSheet
            ->getStyle('E1:E200')
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('ffffff');

        $richTextPersonnel = new RichText();
        $personnelTexte = $richTextPersonnel->createTextRun('1- Renseigner les personnels de l\'unité : ');
        $personnelTexte->getFont()->setBold(true);
        $personnelTexte->getFont()->setName('Century Gothic');
        $richTextPersonnel->createText('copiez-collez ci-dessous dans le tableau de gauche les noms, prénoms et noms des équipes internes depuis l\'onglet 2.1 "RH-' ."\n".
 'personnels" du tableau des données de caractérisation et de production.');

        $macroSheet->setCellValue('B5', $richTextPersonnel);
        $macroSheet->getStyle('B5')->getAlignment()->setWrapText(true);
        $macroSheet->getStyle('B5')->getAlignment()->setVertical('center');

        $richTextDoctorant = new RichText();
        $doctorantTexte = $richTextDoctorant->createTextRun('2- Renseigner les doctorants : ');
        $doctorantTexte->getFont()->setBold(true);
        $doctorantTexte->getFont()->setName('Century Gothic');
        $richTextDoctorant->createText('copiez-collez ci-dessous dans le tableau de droite les noms, prénoms et noms de l\'équipe interne depuis l\'onglet 2.2 "RH-doctorants" du' ."\n".
'tableau des données de caractérisation et de production.');

        $macroSheet->setCellValue('B6', $richTextDoctorant);
        $macroSheet->getStyle('B6')->getAlignment()->setWrapText(true);
        $macroSheet->getStyle('B6')->getAlignment()->setVertical('center');

        $richTextAuto = new RichText();
        $autoTexte = $richTextAuto->createTextRun('3- Remplir automatiquement les colonnes Équipes et Doctorants : ');
        $autoTexte->getFont()->setBold(true);
        $autoTexte->getFont()->setName('Century Gothic');
        $richTextAuto->createText('
cliquez sur le bouton ci-dessous.');

        $macroSheet->setCellValue('B7', $richTextAuto);
        $macroSheet->getStyle('B7')->getAlignment()->setVertical('center');

        $richTextNote = new RichText();
        $noteTexte = $richTextNote->createTextRun('Note : ');
        $noteTexte->getFont()->setBold(true);
        $noteTexte->getFont()->setName('Century Gothic');
        $richTextNote->createText('
en fonction de la quantité de données à traiter, le traitement prend de quelques dizaines de secondes à quelques minutes.');

        $macroSheet->setCellValue('B9', $richTextNote);

        $richTextListeDoc = new RichText();
        $richTextListeDoc->createText('Liste nominative des ');
        $listeDocTexte = $richTextListeDoc->createTextRun('doctorants');
        $listeDocTexte->getFont()->setBold(true);
        $listeDocTexte->getFont()->setName('Century Gothic');
        $richTextListeDoc->createText(' de l\'unité');

        $richTextListePer = new RichText();
        $richTextListePer->createText('Liste nominative des ');
        $listePerTexte = $richTextListePer->createTextRun('personnels');
        $listePerTexte->getFont()->setBold(true);
        $listePerTexte->getFont()->setName('Century Gothic');
        $richTextListePer->createText(' de l\'unité');

        $macroSheet->setCellValue('B12', $richTextListePer);
        $macroSheet->getStyle('B12')->getFont()->setSize(12);
        $macroSheet->getStyle('B12')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('B12')->getAlignment()->setHorizontal('center');

        $macroSheet->setCellValue('F12', $richTextListeDoc);
        $macroSheet->getStyle('F12')->getFont()->setSize(12);
        $macroSheet->getStyle('F12')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('F12')->getAlignment()->setHorizontal('center');

        $macroSheet->setCellValue('B13','Nom');
        $macroSheet->getStyle('B13')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('B13')->getAlignment()->setHorizontal('center');

        $macroSheet->setCellValue('C13','Prénom');
        $macroSheet->getStyle('C13')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('C13')->getAlignment()->setHorizontal('center');

        $macroSheet->setCellValue('D13','Nom de l\'équipe interne');
        $macroSheet->getStyle('D13')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('D13')->getAlignment()->setHorizontal('center');

        $macroSheet
            ->getStyle('B12:D13')
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('ffd9d9d9');

        $macroSheet
            ->getStyle('B12:D12')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');
        $macroSheet
            ->getStyle('B13')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');
        $macroSheet
            ->getStyle('C13')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');
        $macroSheet
            ->getStyle('D13')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');

        $macroSheet->setCellValue('F13','Nom');
        $macroSheet->getStyle('F13')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('F13')->getAlignment()->setHorizontal('center');

        $macroSheet->setCellValue('G13','Prénom');
        $macroSheet->getStyle('G13')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('G13')->getAlignment()->setHorizontal('center');

        $macroSheet->setCellValue('H13','Nom de l\'équipe interne');
        $macroSheet->getStyle('H13')->getAlignment()->setVertical('center');
        $macroSheet->getStyle('H13')->getAlignment()->setHorizontal('center');

        $macroSheet
            ->getStyle('F12:H13')
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB('ffd9d9d9');

        $macroSheet
            ->getStyle('F12:H12')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');
        $macroSheet
            ->getStyle('F13')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');
        $macroSheet
            ->getStyle('G13')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');
        $macroSheet
            ->getStyle('H13')
            ->getBorders()
            ->getOutline()
            ->setBorderStyle(Border::BORDER_THIN)
            ->getColor()
            ->setARGB('ff000000');

        // Déclaration des styles
        $aStyleHeader = ['font' => ['color' => ['argb' => 'FFFFFFFF']],'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER,'vertical' => Alignment::VERTICAL_CENTER,],'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN,'color' => ['argb' => 'FFFFFFFF'],],],'fill' => ['fillType' => Fill::FILL_SOLID,'startColor' => ['argb' => 'ff60497a',],],];
        $aStyleTitle = ['font' => ['bold' => true, 'color' => ['argb' => 'ff000000']],'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER,'vertical' => Alignment::VERTICAL_CENTER,],'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_NONE,],],'fill' => ['fillType' => Fill::FILL_SOLID,'startColor' => ['argb' => 'ffccc0da',],],];
        $aStyleNormal = ['font' => ['color' => ['argb' => 'FF000000']],'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT,'vertical' => Alignment::VERTICAL_CENTER,],'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_NONE,],],'fill' => ['fillType' => Fill::FILL_SOLID,'startColor' => ['argb' => 'FFFFFFFF',],],];

        $richTextEquipe = new RichText();
        $equipeTexte = $richTextEquipe->createTextRun('Equipes');
        $equipeTexte->getFont()->setBold(true);
        $equipeTexte->getFont()->setName('Century Gothic');
        $richTextEquipe->createText('
Indiquer les noms des équipes des auteurs membres de l\'unité, séparés par des points-virgules');

        $richTextDoc = new RichText();
        $doctorant = $richTextDoc->createTextRun('Doctorants');
        $doctorant->getFont()->setBold(true);
        $doctorant->getFont()->setName('Century Gothic');
        $richTextDoc->createText('
Indiquer les noms des doctorants membres de l\'unité et auteurs, séparés par des points-virgules');

        $tab = array();
        // Ventilation par type de document
        foreach ($notices as $notice) {
            switch($notice['docType_s']){
                case 'ART':		  $articles[]  = $notice;		break;
                case 'COMM':      $comms[]     = $notice;		break;
                case 'POSTER':    $posters[]   = $notice;		break;
                case 'PROCEEDINGS' : $procs[]  = $notice;       break;
                case 'ISSUE' :    $issues[]    = $notice;       break;
                case 'OUV':		  $ouvrages[]  = $notice;		break;
                case 'COUV':      $chapitres[] = $notice;		break;
                case 'BLOG' :     $blogs[]     = $notice;       break;
                case 'NOTICE' :   $nots[]      = $notice;       break;
                case 'TRAD' :     $trads[]     = $notice;       break;
                case 'PATENT':    $brevets[]   = $notice;		break;
                case 'OTHER':	  $others[]    = $notice;		break;
                case 'UNDEFINED': $prepublis[] = $notice;		break;
                case 'REPORT':    $rapports[]  = $notice;       break;
                case 'THESE':	  $theses[]    = $notice;		break;
                case 'HDR':		  $hdrs[]      = $notice;		break;
                case 'LECTURE':   $cours[]     = $notice;		break;
                case 'IMG':       $images[]    = $notice;       break;
                case 'VIDEO':     $videos[]    = $notice;       break;
                case 'SON':       $sons[]      = $notice;       break;
                case 'MAP':       $cartes[]    = $notice;       break;
                case 'SOFTWARE':  $logs[]      = $notice;		break;
            }
        }


        // ***********************************************************//
        // Remplissage du document Excel
        //
        ////////////////////////////////////////////////////////////////
        // ****************** ARTICLES *******************************//

        // Création du premier onglet

        $oSheet = $oSpreadsheet->createSheet();

        $oSheet->getTabColor()->setARGB('ffb1a0c7');
        $oSheet->setTitle("Article dans une revue");

        $oSheet->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet->getStyle('A1:P1')->applyFromArray($aStyleHeader);
        $oSheet->getStyle('A1:P1')->getAlignment()->setWrapText(true);

        $oSheet->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet->mergeCells('A1:E1');

        $oSheet->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet->mergeCells('F1:I1');

        $oSheet->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet->mergeCells('J1:L1');

        $oSheet->setCellValue('M1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet->mergeCells('M1:P1');

        $oSheet->getRowDimension(2)->setRowHeight(90);
        $oSheet->getStyle('A2:P2')->applyFromArray($aStyleTitle);
        $oSheet->setAutoFilter('A2:P2');
        $oSheet->getStyle('A2:P2')->getAlignment()->setWrapText(true);
        $oSheet->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet->setCellValue('B2', 'Titre de l\'article');    // colonne 1
        $oSheet->setCellValue('C2', 'Nom de la revue');    // colonne 2
        $oSheet->setCellValue('D2', 'Année');    // colonne 3
        $oSheet->setCellValue('E2', 'Revue à comité de lecture');    // colonne 5
        $oSheet->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet->setCellValue('K2', 'Article en "open access"');    // colonne 11
        $oSheet->setCellValue('L2', 'N° du volume');    // colonne 15
        $oSheet->setCellValue('M2', 'Pages');    // colonne 17
        $oSheet->setCellValue('N2', 'Citation de l\'article');    // colonne 18
        $oSheet->setCellValue('O2', 'Lien DOI');    // colonne 19
        $oSheet->setCellValue('P2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($articles)) {
            foreach ($articles as $article) {
                $oSheet->getRowDimension($ligne)->setRowHeight(35);
                $oSheet->getStyle('A' . $ligne . ':P' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet->getStyle('A' . $ligne . ':P' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($article, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $article)) {
                    $oSheet->setCellValueByColumnAndRow(2, $ligne, $article['title_s'][0]);
                }
                if (array_key_exists("journalTitle_s", $article)) {
                    $oSheet->setCellValueByColumnAndRow(3, $ligne, $article['journalTitle_s']);
                }
                if (array_key_exists("producedDateY_i", $article)) {
                    $oSheet->setCellValueByColumnAndRow(4, $ligne, $article['producedDateY_i']);
                }
                if (array_key_exists("peerReviewing_s", $article)) {
                    if($article['peerReviewing_s']) {
                        $oSheet->setCellValueByColumnAndRow(5, $ligne, 'O');
                    }else{
                        $oSheet->setCellValueByColumnAndRow(5, $ligne, 'N');
                    }
                }
                $oSheet->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($article));
                $oSheet->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($article,$equipes));
                $oSheet->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($article));
                if (array_key_exists("openAccess_bool", $article)) {
                    if ($article['openAccess_bool']) {
                        $oSheet->setCellValueByColumnAndRow(11, $ligne, 'O');
                    } else {
                        $oSheet->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                if (array_key_exists("volume_s", $article)) {
                    $oSheet->setCellValueByColumnAndRow(12, $ligne, $article['volume_s']);
                }
                if (array_key_exists("page_s", $article)) {
                    $oSheet->setCellValueByColumnAndRow(13, $ligne, $article['page_s']);
                }
                $oSheet->setCellValueByColumnAndRow( 14,  $ligne, getCitationArticle($article));
                if (array_key_exists("doiId_s", $article)) {
                    $oSheet->setCellValueByColumnAndRow(15, $ligne, 'http://doi.org/' . $article['doiId_s']);
                    $oSheet->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $article['doiId_s']);
                }
                if (array_key_exists("halId_s", $article)) {
                    $oSheet->setCellValueByColumnAndRow(16, $ligne, 'https://hal.archives-ouvertes.fr/' . $article['halId_s']);
                    $oSheet->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $article['halId_s']);
                }

                $oSheet->getStyle('A'.$ligne.':P'.$ligne)->getAlignment()->setVertical('top');
                $oSheet->getStyle('D'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet->getStyle('L'.$ligne.':M'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** COMM *******************************//

        $oSheet2 = $oSpreadsheet->createSheet();

        $oSheet2->getTabColor()->setARGB('ffb1a0c7');
        $oSheet2->setTitle("Communication dans un congres");

        $oSheet2->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet2->getStyle('A1:S1')->applyFromArray($aStyleHeader);
        $oSheet2->getStyle('A1:S1')->getAlignment()->setWrapText(true);

        $oSheet2->setCellValue('A1', 'Description du congrès');    // colonne 0
        $oSheet2->mergeCells('A1:I1');

        $oSheet2->setCellValue('J1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet2->mergeCells('J1:M1');

        $oSheet2->setCellValue('N1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet2->mergeCells('N1:P1');

        $oSheet2->setCellValue('Q1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet2->mergeCells('Q1:S1');

        $oSheet2->getRowDimension(2)->setRowHeight(90);
        $oSheet2->getStyle('A2:S2')->applyFromArray($aStyleTitle);
        $oSheet2->setAutoFilter('A2:S2');
        $oSheet2->getStyle('A2:S2')->getAlignment()->setWrapText(true);
        $oSheet2->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet2->setCellValue('B2', 'Titre de la communication');    // colonne 1
        $oSheet2->setCellValue('C2', 'Titre du congrès');    // colonne 2
        $oSheet2->setCellValue('D2', 'Ville du congrès');    // colonne 3
        $oSheet2->setCellValue('E2', 'Pays du congrès');    // colonne 4
        $oSheet2->setCellValue('F2', 'Année');    // colonne 5
        $oSheet2->setCellValue('G2', 'Date de début du congrès');    // colonne 6
        $oSheet2->setCellValue('H2', 'Date de fin du congrès');    // colonne 7
        $oSheet2->setCellValue('I2', 'Conférence invitée');    // colonne 8
        $oSheet2->setCellValue('J2', $richTextEquipe);    // colonne 9
        $oSheet2->setCellValue('K2', $richTextDoc);    // colonne 10
        $oSheet2->setCellValue('L2', 'Affiliation institutionnelle des co-auteurs');    // colonne 11
        $oSheet2->setCellValue('M2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 12
        $oSheet2->setCellValue('N2', 'Audience');    // colonne 13
        $oSheet2->setCellValue('O2', 'Langue');    // colonne 14
        $oSheet2->setCellValue('P2', 'Article en "open access"');    // colonne 15
        $oSheet2->setCellValue('Q2', 'Citation des actes');    // colonne 20
        $oSheet2->setCellValue('R2', 'Lien DOI');    // colonne 21
        $oSheet2->setCellValue('S2', 'Lien HAL');    // colonne 22

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet2->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('Q')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('R')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet2->getColumnDimension('S')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($comms)) {
            foreach ($comms as $comm) {
                $oSheet2->getRowDimension($ligne)->setRowHeight(35);
                $oSheet2->getStyle('A' . $ligne . ':S' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet2->getStyle('A' . $ligne . ':S' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet2->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($comm, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(2, $ligne, $comm['title_s'][0]);
                }
                if (array_key_exists("conferenceTitle_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(3, $ligne, $comm['conferenceTitle_s']);
                }
                if (array_key_exists("city_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(4, $ligne, $comm['city_s']);
                }
                if (array_key_exists("country_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(5, $ligne, codeToCountry($comm['country_s']));
                }
                if (array_key_exists("producedDateY_i", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(6, $ligne, $comm['producedDateY_i']);
                }
                if (array_key_exists("conferenceStartDate_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(7, $ligne, $comm['conferenceStartDate_s']);
                }
                if (array_key_exists("conferenceEndDate_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(8, $ligne, $comm['conferenceEndDate_s']);
                }
                if (array_key_exists("invitedCommunication_s", $comm)) {
                    if($comm['invitedCommunication_s']) {
                        $oSheet2->setCellValueByColumnAndRow(9, $ligne, 'O');
                    }else{
                        $oSheet2->setCellValueByColumnAndRow(9, $ligne, 'N');
                    }
                }
                $oSheet2->setCellValueByColumnAndRow(10, $ligne, '');
                $oSheet2->setCellValueByColumnAndRow(11, $ligne, '');
                $oSheet2->setCellValueByColumnAndRow(12, $ligne, getCoAuteursHceres($comm));
                $oSheet2->setCellValueByColumnAndRow( 13,  $ligne, getPremierDernier($comm,$equipes));
                $oSheet2->setCellValueByColumnAndRow(14, $ligne, getAudience($comm));
                $oSheet2->setCellValueByColumnAndRow(15, $ligne, getLangueHceres($comm));
                if (array_key_exists("openAccess_bool", $comm)) {
                    if($comm['openAccess_bool']) {
                        $oSheet2->setCellValueByColumnAndRow(16, $ligne, 'O');
                    }else{
                        $oSheet2->setCellValueByColumnAndRow(16, $ligne, 'N');
                    }
                }
                $oSheet2->setCellValueByColumnAndRow( 17,  $ligne, getCitationCommunication($comm));
                if (array_key_exists("doiId_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(18, $ligne, 'http://doi.org/' . $comm['doiId_s']);
                    $oSheet2->getCellByColumnAndRow(18, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $comm['doiId_s']);
                }
                if (array_key_exists("halId_s", $comm)) {
                    $oSheet2->setCellValueByColumnAndRow(19, $ligne, 'https://hal.archives-ouvertes.fr/' . $comm['halId_s']);
                    $oSheet2->getCellByColumnAndRow(19, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $comm['halId_s']);
                }

                $oSheet2->getStyle('A'.$ligne.':S'.$ligne)->getAlignment()->setVertical('top');
                $oSheet2->getStyle('F'.$ligne.':P'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet2->getStyle('L'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet2->getStyle('A' . $ligne . ':S' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet2->getStyle('A' . $ligne . ':S' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** POSTER *****************************//

        $oSheet3 = $oSpreadsheet->createSheet();

        $oSheet3->getTabColor()->setARGB('ffb1a0c7');
        $oSheet3->setTitle("Poster");

        $oSheet3->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet3->getStyle('A1:S1')->applyFromArray($aStyleHeader);
        $oSheet3->getStyle('A1:S1')->getAlignment()->setWrapText(true);

        $oSheet3->setCellValue('A1', 'Description du congrès');    // colonne 0
        $oSheet3->mergeCells('A1:I1');

        $oSheet3->setCellValue('J1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet3->mergeCells('J1:M1');

        $oSheet3->setCellValue('N1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet3->mergeCells('N1:P1');

        $oSheet3->setCellValue('Q1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet3->mergeCells('Q1:S1');

        $oSheet3->getRowDimension(2)->setRowHeight(90);
        $oSheet3->getStyle('A2:S2')->applyFromArray($aStyleTitle);
        $oSheet3->setAutoFilter('A2:S2');
        $oSheet3->getStyle('A2:S2')->getAlignment()->setWrapText(true);
        $oSheet3->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet3->setCellValue('B2', 'Titre du poster');    // colonne 1
        $oSheet3->setCellValue('C2', 'Conférence invitée');    // colonne 2
        $oSheet3->setCellValue('D2', 'Titre du congrès');    // colonne 3
        $oSheet3->setCellValue('E2', 'Ville du congrès');    // colonne 4
        $oSheet3->setCellValue('F2', 'Pays du congrès');    // colonne 5
        $oSheet3->setCellValue('G2', 'Année');    // colonne 6
        $oSheet3->setCellValue('H2', 'Date de début du congrès');    // colonne 7
        $oSheet3->setCellValue('I2', 'Date de fin du congrès');    // colonne 8
        $oSheet3->setCellValue('J2', $richTextEquipe);    // colonne 9
        $oSheet3->setCellValue('K2', $richTextDoc);    // colonne 10
        $oSheet3->setCellValue('L2', 'Affiliation institutionnelle des co-auteurs');    // colonne 11
        $oSheet3->setCellValue('M2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 12
        $oSheet3->setCellValue('N2', 'Audience');    // colonne 13
        $oSheet3->setCellValue('O2', 'Langue');    // colonne 14
        $oSheet3->setCellValue('P2', 'Article en "open access"');    // colonne 15
        $oSheet3->setCellValue('Q2', 'Citation des actes');    // colonne 20
        $oSheet3->setCellValue('R2', 'Lien DOI');    // colonne 21
        $oSheet3->setCellValue('S2', 'Lien HAL');    // colonne 22

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet3->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('D')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('Q')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('R')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet3->getColumnDimension('S')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($posters)) {
            foreach ($posters as $poster) {
                $oSheet3->getRowDimension($ligne)->setRowHeight(35);
                $oSheet3->getStyle('A' . $ligne . ':S' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet3->getStyle('A' . $ligne . ':S' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet3->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($poster, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(2, $ligne, $poster['title_s'][0]);
                }
                if (array_key_exists("invitedCommunication_s", $poster)) {
                    if($poster['invitedCommunication_s']) {
                        $oSheet3->setCellValueByColumnAndRow(3, $ligne, 'O');
                    }else{
                        $oSheet3->setCellValueByColumnAndRow(3, $ligne, 'N');
                    }
                }
                if (array_key_exists("conferenceTitle_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(4, $ligne, $poster['conferenceTitle_s']);
                }
                if (array_key_exists("city_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(5, $ligne, $poster['city_s']);
                }
                if (array_key_exists("country_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(6, $ligne, codeToCountry($poster['country_s']));
                }
                if (array_key_exists("producedDateY_i", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(7, $ligne, $poster['producedDateY_i']);
                }
                if (array_key_exists("conferenceStartDate_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(8, $ligne, $poster['conferenceStartDate_s']);
                }
                if (array_key_exists("conferenceEndDate_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(9, $ligne, $poster['conferenceEndDate_s']);
                }
                $oSheet3->setCellValueByColumnAndRow(10, $ligne, '');
                $oSheet3->setCellValueByColumnAndRow(11, $ligne, '');
                $oSheet3->setCellValueByColumnAndRow(12, $ligne, getCoAuteursHceres($poster));
                $oSheet3->setCellValueByColumnAndRow( 13,  $ligne, getPremierDernier($poster,$equipes));
                $oSheet3->setCellValueByColumnAndRow(14, $ligne, getAudience($poster));
                $oSheet3->setCellValueByColumnAndRow(15, $ligne, getLangueHceres($poster));
                if (array_key_exists("openAccess_bool", $poster)) {
                    if($poster['openAccess_bool']) {
                        $oSheet3->setCellValueByColumnAndRow(16, $ligne, 'O');
                    }else{
                        $oSheet3->setCellValueByColumnAndRow(16, $ligne, 'N');
                    }
                }
                $oSheet3->setCellValueByColumnAndRow( 17,  $ligne, getCitationCommunication($poster));
                if (array_key_exists("doiId_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(18, $ligne, 'http://doi.org/' . $poster['doiId_s']);
                    $oSheet3->getCellByColumnAndRow(18, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $poster['doiId_s']);
                }
                if (array_key_exists("halId_s", $poster)) {
                    $oSheet3->setCellValueByColumnAndRow(19, $ligne, 'https://hal.archives-ouvertes.fr/' . $poster['halId_s']);
                    $oSheet3->getCellByColumnAndRow(19, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $poster['halId_s']);
                }

                $oSheet3->getStyle('A'.$ligne.':S'.$ligne)->getAlignment()->setVertical('top');
                $oSheet3->getStyle('G'.$ligne.':P'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet3->getStyle('C'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet3->getStyle('L'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet3->getStyle('A' . $ligne . ':S' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet3->getStyle('A' . $ligne . ':S' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** PROCEEDINGS ************************//

        $oSheet4 = $oSpreadsheet->createSheet();

        $oSheet4->getTabColor()->setARGB('ffb1a0c7');
        $oSheet4->setTitle("Proceedings Recueil des comm.");

        $oSheet4->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet4->getStyle('A1:P1')->applyFromArray($aStyleHeader);
        $oSheet4->getStyle('A1:P1')->getAlignment()->setWrapText(true);

        $oSheet4->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet4->mergeCells('A1:D1');

        $oSheet4->setCellValue('E1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet4->mergeCells('E1:H1');

        $oSheet4->setCellValue('I1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet4->mergeCells('I1:J1');

        $oSheet4->setCellValue('K1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet4->mergeCells('K1:P1');

        $oSheet4->getRowDimension(2)->setRowHeight(90);
        $oSheet4->getStyle('A2:P2')->applyFromArray($aStyleTitle);
        $oSheet4->setAutoFilter('A2:P2');
        $oSheet4->getStyle('A2:P2')->getAlignment()->setWrapText(true);
        $oSheet4->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet4->setCellValue('B2', 'Titre de la publication');    // colonne 1
        $oSheet4->setCellValue('C2', 'Titre du volume');    // colonne 1
        $oSheet4->setCellValue('D2', 'Année');    // colonne 3
        $oSheet4->setCellValue('E2', $richTextEquipe);    // colonne 6
        $oSheet4->setCellValue('F2', $richTextDoc);    // colonne 7
        $oSheet4->setCellValue('G2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet4->setCellValue('H2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet4->setCellValue('I2', 'Langue');    // colonne 10
        $oSheet4->setCellValue('J2', 'Article en "open access"');    // colonne 11
        $oSheet4->setCellValue('K2', 'N° du volume');    // colonne 15
        $oSheet4->setCellValue('L2', 'N°');    // colonne 15
        $oSheet4->setCellValue('M2', 'Pages');    // colonne 17
        $oSheet4->setCellValue('N2', 'Citation de l\'article');    // colonne 18
        $oSheet4->setCellValue('O2', 'Lien DOI');    // colonne 19
        $oSheet4->setCellValue('P2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet4->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('G')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet4->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($procs)) {
            foreach ($procs as $proc) {
                $oSheet4->getRowDimension($ligne)->setRowHeight(35);
                $oSheet4->getStyle('A' . $ligne . ':P' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet4->getStyle('A' . $ligne . ':P' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet4->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($proc, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(2, $ligne, $proc['title_s'][0]);
                }
                if (array_key_exists("serie_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(3, $ligne, $proc['serie_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(4, $ligne, $proc['producedDateY_i']);
                }
                $oSheet4->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet4->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet4->setCellValueByColumnAndRow(7, $ligne, getCoAuteursHceres($proc));
                $oSheet4->setCellValueByColumnAndRow( 8,  $ligne, getPremierDernier($proc,$equipes));
                $oSheet4->setCellValueByColumnAndRow(9, $ligne, getLangueHceres($proc));
                if (array_key_exists("openAccess_bool", $proc)) {
                    if ($proc['openAccess_bool']) {
                        $oSheet4->setCellValueByColumnAndRow(10, $ligne, 'O');
                    } else {
                        $oSheet4->setCellValueByColumnAndRow(10, $ligne, 'N');
                    }
                }
                if (array_key_exists("volume_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(11, $ligne, $proc['volume_s']);
                }
                if (array_key_exists("number_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(12, $ligne, $proc['number_s'][0]);
                }
                if (array_key_exists("page_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(13, $ligne, $proc['page_s']);
                }
                $oSheet4->setCellValueByColumnAndRow( 14,  $ligne, getCitationAutre($proc));
                if (array_key_exists("doiId_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(15, $ligne, 'http://doi.org/' . $proc['doiId_s']);
                    $oSheet4->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $proc['doiId_s']);
                }
                if (array_key_exists("halId_s", $proc)) {
                    $oSheet4->setCellValueByColumnAndRow(16, $ligne, 'https://hal.archives-ouvertes.fr/' . $proc['halId_s']);
                    $oSheet4->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $proc['halId_s']);
                }

                $oSheet4->getStyle('A'.$ligne.':P'.$ligne)->getAlignment()->setVertical('top');
                $oSheet4->getStyle('D'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet4->getStyle('G'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet4->getStyle('K'.$ligne.':M'.$ligne)->getAlignment()->setHorizontal('right');
                $oSheet4->getStyle('N'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet4->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet4->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** ISSUES *****************************//

        $oSheet5 = $oSpreadsheet->createSheet();

        $oSheet5->getTabColor()->setARGB('ffb1a0c7');
        $oSheet5->setTitle("No special de revue");

        $oSheet5->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet5->getStyle('A1:O1')->applyFromArray($aStyleHeader);
        $oSheet5->getStyle('A1:O1')->getAlignment()->setWrapText(true);

        $oSheet5->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet5->mergeCells('A1:D1');

        $oSheet5->setCellValue('E1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet5->mergeCells('E1:H1');

        $oSheet5->setCellValue('I1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet5->mergeCells('I1:J1');

        $oSheet5->setCellValue('K1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet5->mergeCells('K1:O1');

        $oSheet5->getRowDimension(2)->setRowHeight(90);
        $oSheet5->getStyle('A2:O2')->applyFromArray($aStyleTitle);
        $oSheet5->setAutoFilter('A2:O2');
        $oSheet5->getStyle('A2:O2')->getAlignment()->setWrapText(true);
        $oSheet5->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet5->setCellValue('B2', 'Titre de la publication');    // colonne 1
        $oSheet5->setCellValue('C2', 'Titre du volume');    // colonne 1
        $oSheet5->setCellValue('D2', 'Année');    // colonne 3
        $oSheet5->setCellValue('E2', $richTextEquipe);    // colonne 6
        $oSheet5->setCellValue('F2', $richTextDoc);    // colonne 7
        $oSheet5->setCellValue('G2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet5->setCellValue('H2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet5->setCellValue('I2', 'Langue');    // colonne 10
        $oSheet5->setCellValue('J2', 'Article en "open access"');    // colonne 11
        $oSheet5->setCellValue('K2', 'Citation de l\'article');    // colonne 18
        $oSheet5->setCellValue('L2', 'N° du volume');    // colonne 15
        $oSheet5->setCellValue('M2', 'N°');    // colonne 15
        $oSheet5->setCellValue('N2', 'Lien DOI');    // colonne 19
        $oSheet5->setCellValue('O2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet5->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('G')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('K')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet5->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($issues)) {
            foreach ($issues as $issue) {
                $oSheet5->getRowDimension($ligne)->setRowHeight(35);
                $oSheet5->getStyle('A' . $ligne . ':O' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet5->getStyle('A' . $ligne . ':O' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet5->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($issue, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(2, $ligne, $issue['title_s'][0]);
                }
                if (array_key_exists("serie_s", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(3, $ligne, $issue['serie_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(4, $ligne, $issue['producedDateY_i']);
                }
                $oSheet5->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet5->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet5->setCellValueByColumnAndRow(7, $ligne, getCoAuteursHceres($issue));
                $oSheet5->setCellValueByColumnAndRow( 8,  $ligne, getPremierDernier($issue,$equipes));
                $oSheet5->setCellValueByColumnAndRow(9, $ligne, getLangueHceres($issue));
                if (array_key_exists("openAccess_bool", $issue)) {
                    if ($issue['openAccess_bool']) {
                        $oSheet5->setCellValueByColumnAndRow(10, $ligne, 'O');
                    } else {
                        $oSheet5->setCellValueByColumnAndRow(10, $ligne, 'N');
                    }
                }
                $oSheet5->setCellValueByColumnAndRow( 11,  $ligne, getCitationAutre($issue));
                if (array_key_exists("volume_s", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(12, $ligne, $issue['volume_s']);
                }
                if (array_key_exists("number_s", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(13, $ligne, $issue['number_s'][0]);
                }
                if (array_key_exists("doiId_s", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(14, $ligne, 'http://doi.org/' . $issue['doiId_s']);
                    $oSheet5->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $issue['doiId_s']);
                }
                if (array_key_exists("halId_s", $issue)) {
                    $oSheet5->setCellValueByColumnAndRow(15, $ligne, 'https://hal.archives-ouvertes.fr/' . $issue['halId_s']);
                    $oSheet5->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $issue['halId_s']);
                }

                $oSheet5->getStyle('A'.$ligne.':O'.$ligne)->getAlignment()->setVertical('top');
                $oSheet5->getStyle('D'.$ligne.':J'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet5->getStyle('G'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet5->getStyle('L'.$ligne.':M'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet5->getStyle('A' . $ligne . ':O' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet5->getStyle('A' . $ligne . ':O' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** OUVRAGE ****************************//

        $oSheet6 = $oSpreadsheet->createSheet();

        $oSheet6->getTabColor()->setARGB('ffb1a0c7');
        $oSheet6->setTitle("Ouvrage");

        $oSheet6->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet6->getStyle('A1:Q1')->applyFromArray($aStyleHeader);
        $oSheet6->getStyle('A1:Q1')->getAlignment()->setWrapText(true);

        $oSheet6->setCellValue('A1', 'Description de l\'ouvrage');    // colonne 0
        $oSheet6->mergeCells('A1:G1');

        $oSheet6->setCellValue('H1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet6->mergeCells('H1:K1');

        $oSheet6->setCellValue('L1', 'Caractéristiques de Fl\'article');    // colonne 0
        $oSheet6->mergeCells('L1:M1');

        $oSheet6->setCellValue('N1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet6->mergeCells('N1:Q1');

        $oSheet6->getRowDimension(2)->setRowHeight(90);
        $oSheet6->getStyle('A2:Q2')->applyFromArray($aStyleTitle);
        $oSheet6->setAutoFilter('A2:Q2');
        $oSheet6->getStyle('A2:Q2')->getAlignment()->setWrapText(true);
        $oSheet6->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet6->setCellValue('B2', 'Titre de l\'ouvrage');    // colonne 1
        $oSheet6->setCellValue('C2', 'Sous-titre');    // colonne 2
        $oSheet6->setCellValue('D2', 'Résumé');    // colonne 3
        $oSheet6->setCellValue('E2', 'Mots clés');    // colonne 4
        $oSheet6->setCellValue('F2', 'Editeur');    // colonne 5
        $oSheet6->setCellValue('G2', 'Année');    // colonne 6
        $oSheet6->setCellValue('H2', $richTextEquipe);    // colonne 9
        $oSheet6->setCellValue('I2', $richTextDoc);    // colonne 10
        $oSheet6->setCellValue('J2', 'Affiliation institutionnelle des co-auteurs');    // colonne 11
        $oSheet6->setCellValue('K2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 12
        $oSheet6->setCellValue('L2', 'Langue');    // colonne 14
        $oSheet6->setCellValue('M2', 'Ouvrage en "open access"');    // colonne 15
        $oSheet6->setCellValue('N2', 'Pages');    // colonne 19
        $oSheet6->setCellValue('O2', 'Citation de l\'article');    // colonne 20
        $oSheet6->setCellValue('P2', 'Lien DOI');    // colonne 21
        $oSheet6->setCellValue('Q2', 'Lien HAL');    // colonne 22

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet6->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('D')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('E')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('F')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('J')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('O')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet6->getColumnDimension('Q')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($ouvrages)) {
            foreach ($ouvrages as $ouvrage) {
                $oSheet6->getRowDimension($ligne)->setRowHeight(35);
                $oSheet6->getStyle('A' . $ligne . ':Q' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet6->getStyle('A' . $ligne . ':Q' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet6->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($ouvrage, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(2, $ligne, $ouvrage['title_s'][0]);
                }
                if (array_key_exists("subTitle_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(3, $ligne, $ouvrage['subTitle_s'][0]);
                }
                if (array_key_exists("abstract_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(4, $ligne, $ouvrage['abstract_s'][0]);
                }
                $oSheet6->setCellValueByColumnAndRow(5, $ligne, getKeyword($ouvrage));
                if (array_key_exists("publisher_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(6, $ligne, $ouvrage['publisher_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(7, $ligne, $ouvrage['producedDateY_i']);
                }
                $oSheet6->setCellValueByColumnAndRow(8, $ligne, '');
                $oSheet6->setCellValueByColumnAndRow(9, $ligne, '');
                $oSheet6->setCellValueByColumnAndRow(10, $ligne, getCoAuteursHceres($ouvrage));
                $oSheet6->setCellValueByColumnAndRow( 11,  $ligne, getPremierDernier($ouvrage,$equipes));
                $oSheet6->setCellValueByColumnAndRow(12, $ligne, getLangueHceres($ouvrage));
                if (array_key_exists("openAccess_bool", $ouvrage)) {
                    if($ouvrage['openAccess_bool']) {
                        $oSheet6->setCellValueByColumnAndRow(13, $ligne, 'O');
                    }else{
                        $oSheet6->setCellValueByColumnAndRow(13, $ligne, 'N');
                    }
                }
                if (array_key_exists("page_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(14, $ligne, $ouvrage['page_s']);
                }
                $oSheet6->setCellValueByColumnAndRow( 15,  $ligne, getCitationOuvrage($ouvrage));
                if (array_key_exists("doiId_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(16, $ligne, 'http://doi.org/' . $ouvrage['doiId_s']);
                    $oSheet6->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $ouvrage['doiId_s']);
                }
                if (array_key_exists("halId_s", $ouvrage)) {
                    $oSheet6->setCellValueByColumnAndRow(17, $ligne, 'https://hal.archives-ouvertes.fr/' . $ouvrage['halId_s']);
                    $oSheet6->getCellByColumnAndRow(17, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $ouvrage['halId_s']);
                }

                $oSheet6->getStyle('A'.$ligne.':Q'.$ligne)->getAlignment()->setVertical('top');
                $oSheet6->getStyle('G'.$ligne.':M'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet6->getStyle('J'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet6->getStyle('N'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet6->getStyle('A' . $ligne . ':Q' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet6->getStyle('A' . $ligne . ':Q' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** CHAPITRE OUVRAGE *****************//

        $oSheet7 = $oSpreadsheet->createSheet();

        $oSheet7->getTabColor()->setARGB('ffb1a0c7');
        $oSheet7->setTitle("Chapitre ouvrage");

        $oSheet7->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet7->getStyle('A1:O1')->applyFromArray($aStyleHeader);
        $oSheet7->getStyle('A1:O1')->getAlignment()->setWrapText(true);

        $oSheet7->setCellValue('A1', 'Description du chapitre dʹouvrage');    // colonne 0
        $oSheet7->mergeCells('A1:E1');

        $oSheet7->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet7->mergeCells('F1:I1');

        $oSheet7->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet7->mergeCells('J1:K1');

        $oSheet7->setCellValue('L1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet7->mergeCells('L1:O1');

        $oSheet7->getRowDimension(2)->setRowHeight(90);
        $oSheet7->getStyle('A2:O2')->applyFromArray($aStyleTitle);
        $oSheet7->setAutoFilter('A2:O2');
        $oSheet7->getStyle('A2:O2')->getAlignment()->setWrapText(true);
        $oSheet7->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet7->setCellValue('B2', 'Titre du chapitre dʹouvrage');    // colonne 1
        $oSheet7->setCellValue('C2', 'Titre de l\'ouvrage');    // colonne 2
        $oSheet7->setCellValue('D2', 'Editeur');    // colonne 3
        $oSheet7->setCellValue('E2', 'Année');    // colonne 4
        $oSheet7->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet7->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet7->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet7->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet7->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet7->setCellValue('K2', 'Ouvrage en "open access"');    // colonne 11
        $oSheet7->setCellValue('L2', 'Pages');    // colonne 15
        $oSheet7->setCellValue('M2', 'Citation de l\'article');    // colonne 16
        $oSheet7->setCellValue('N2', 'Lien DOI');    // colonne 17
        $oSheet7->setCellValue('O2', 'Lien HAL');    // colonne 18

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet7->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('D')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('M')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet7->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($chapitres)) {
            foreach ($chapitres as $chapitre) {
                $oSheet7->getRowDimension($ligne)->setRowHeight(35);
                $oSheet7->getStyle('A' . $ligne . ':O' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet7->getStyle('A' . $ligne . ':O' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet7->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($chapitre, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(2, $ligne, $chapitre['title_s'][0]);
                }
                if (array_key_exists("bookTitle_s", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(3, $ligne, $chapitre['bookTitle_s']);
                }
                if (array_key_exists("publisher_s", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(4, $ligne, $chapitre['publisher_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(5, $ligne, $chapitre['producedDateY_i']);
                }
                $oSheet7->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet7->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet7->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($chapitre));
                $oSheet7->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($chapitre,$equipes));
                $oSheet7->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($chapitre));
                if (array_key_exists("openAccess_bool", $chapitre)) {
                    if($chapitre['openAccess_bool']) {
                        $oSheet7->setCellValueByColumnAndRow(11, $ligne, 'O');
                    }else{
                        $oSheet7->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                if (array_key_exists("page_s", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(12, $ligne, $chapitre['page_s']);
                }
                $oSheet7->setCellValueByColumnAndRow( 13,  $ligne, getCitationChapitreOuvrage($chapitre));
                if (array_key_exists("doiId_s", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(14, $ligne, 'http://doi.org/' . $chapitre['doiId_s']);
                    $oSheet7->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $chapitre['doiId_s']);
                }
                if (array_key_exists("halId_s", $chapitre)) {
                    $oSheet7->setCellValueByColumnAndRow(15, $ligne, 'https://hal.archives-ouvertes.fr/' . $chapitre['halId_s']);
                    $oSheet7->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $chapitre['halId_s']);
                }

                $oSheet7->getStyle('A'.$ligne.':S'.$ligne)->getAlignment()->setVertical('top');
                $oSheet7->getStyle('E'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet7->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet7->getStyle('L'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet7->getStyle('A' . $ligne . ':O' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet7->getStyle('A' . $ligne . ':O' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** BLOG *******************************//

        $oSheet8 = $oSpreadsheet->createSheet();

        $oSheet8->getTabColor()->setARGB('ffb1a0c7');
        $oSheet8->setTitle("Article de blog scientifique");

        $oSheet8->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet8->getStyle('A1:P1')->applyFromArray($aStyleHeader);
        $oSheet8->getStyle('A1:P1')->getAlignment()->setWrapText(true);

        $oSheet8->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet8->mergeCells('A1:D1');

        $oSheet8->setCellValue('E1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet8->mergeCells('E1:H1');

        $oSheet8->setCellValue('I1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet8->mergeCells('I1:J1');

        $oSheet8->setCellValue('K1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet8->mergeCells('K1:P1');

        $oSheet8->getRowDimension(2)->setRowHeight(90);
        $oSheet8->getStyle('A2:P2')->applyFromArray($aStyleTitle);
        $oSheet8->setAutoFilter('A2:P2');
        $oSheet8->getStyle('A2:P2')->getAlignment()->setWrapText(true);
        $oSheet8->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet8->setCellValue('B2', 'Titre de la publication');    // colonne 1
        $oSheet8->setCellValue('C2', 'Nom du blog ou du carnet');    // colonne 1
        $oSheet8->setCellValue('D2', 'Année');    // colonne 3
        $oSheet8->setCellValue('E2', $richTextEquipe);    // colonne 6
        $oSheet8->setCellValue('F2', $richTextDoc);    // colonne 7
        $oSheet8->setCellValue('G2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet8->setCellValue('H2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet8->setCellValue('I2', 'Langue');    // colonne 10
        $oSheet8->setCellValue('J2', 'Article en "open access"');    // colonne 11
        $oSheet8->setCellValue('K2', 'Description');    // colonne 14
        $oSheet8->setCellValue('L2', 'Citation de l\'article');    // colonne 18
        $oSheet8->setCellValue('M2', 'Résumé');    // colonne 4
        $oSheet8->setCellValue('N2', 'Mots clés');    // colonne 4
        $oSheet8->setCellValue('O2', 'Lien DOI');    // colonne 19
        $oSheet8->setCellValue('P2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet8->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('G')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('K')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('M')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet8->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($blogs)) {
            foreach ($blogs as $blog) {
                $oSheet8->getRowDimension($ligne)->setRowHeight(35);
                $oSheet8->getStyle('A' . $ligne . ':P' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet8->getStyle('A' . $ligne . ':P' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet8->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($blog, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(2, $ligne, $blog['title_s'][0]);
                }
                if (array_key_exists("journalTitle_s", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(3, $ligne, $blog['journalTitle_s']);
                }
                if (array_key_exists("producedDateY_i", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(4, $ligne, $blog['producedDateY_i']);
                }
                $oSheet8->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet8->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet8->setCellValueByColumnAndRow(7, $ligne, getCoAuteursHceres($blog));
                $oSheet8->setCellValueByColumnAndRow( 8,  $ligne, getPremierDernier($blog,$equipes));
                $oSheet8->setCellValueByColumnAndRow(9, $ligne, getLangueHceres($blog));
                if (array_key_exists("openAccess_bool", $blog)) {
                    if ($blog['openAccess_bool']) {
                        $oSheet8->setCellValueByColumnAndRow(10, $ligne, 'O');
                    } else {
                        $oSheet8->setCellValueByColumnAndRow(10, $ligne, 'N');
                    }
                }
                if (array_key_exists("description_s", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(11, $ligne, $blog['description_s']);
                }
                $oSheet8->setCellValueByColumnAndRow( 12,  $ligne, getCitationAutre($blog));
                if (array_key_exists("abstract_s", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(13, $ligne, $blog['abstract_s'][0]);
                }
                $oSheet8->setCellValueByColumnAndRow(14, $ligne, getKeyword($blog));
                if (array_key_exists("doiId_s", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(15, $ligne, 'http://doi.org/' . $blog['doiId_s']);
                    $oSheet8->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $blog['doiId_s']);
                }
                if (array_key_exists("halId_s", $blog)) {
                    $oSheet8->setCellValueByColumnAndRow(16, $ligne, 'https://hal.archives-ouvertes.fr/' . $blog['halId_s']);
                    $oSheet8->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $blog['halId_s']);
                }

                $oSheet8->getStyle('A'.$ligne.':P'.$ligne)->getAlignment()->setVertical('top');
                $oSheet8->getStyle('D'.$ligne.':J'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet8->getStyle('G'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet8->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet8->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** NOTICE *****************************//

        $oSheet9 = $oSpreadsheet->createSheet();

        $oSheet9->getTabColor()->setARGB('ffb1a0c7');
        $oSheet9->setTitle("Not. encyclopedie dictionnaire");

        $oSheet9->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet9->getStyle('A1:Q1')->applyFromArray($aStyleHeader);
        $oSheet9->getStyle('A1:Q1')->getAlignment()->setWrapText(true);

        $oSheet9->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet9->mergeCells('A1:D1');

        $oSheet9->setCellValue('E1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet9->mergeCells('E1:H1');

        $oSheet9->setCellValue('I1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet9->mergeCells('I1:J1');

        $oSheet9->setCellValue('K1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet9->mergeCells('K1:Q1');

        $oSheet9->getRowDimension(2)->setRowHeight(90);
        $oSheet9->getStyle('A2:Q2')->applyFromArray($aStyleTitle);
        $oSheet9->setAutoFilter('A2:Q2');
        $oSheet9->getStyle('A2:Q2')->getAlignment()->setWrapText(true);
        $oSheet9->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet9->setCellValue('B2', 'Titre de la publication');    // colonne 1
        $oSheet9->setCellValue('C2', 'Nom de la revue');    // colonne 14
        $oSheet9->setCellValue('D2', 'Année');    // colonne 3
        $oSheet9->setCellValue('E2', $richTextEquipe);    // colonne 6
        $oSheet9->setCellValue('F2', $richTextDoc);    // colonne 7
        $oSheet9->setCellValue('G2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet9->setCellValue('H2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet9->setCellValue('I2', 'Langue');    // colonne 10
        $oSheet9->setCellValue('J2', 'Article en "open access"');    // colonne 11
        $oSheet9->setCellValue('K2', 'N° du volume');    // colonne 14
        $oSheet9->setCellValue('L2', 'Pages');    // colonne 17
        $oSheet9->setCellValue('M2', 'Citation de l\'article');    // colonne 18
        $oSheet9->setCellValue('N2', 'Résumé');    // colonne 4
        $oSheet9->setCellValue('O2', 'Mots clés');    // colonne 4
        $oSheet9->setCellValue('P2', 'Lien DOI');    // colonne 19
        $oSheet9->setCellValue('Q2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet9->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('G')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('M')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('O')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet9->getColumnDimension('Q')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($nots)) {
            foreach ($nots as $not) {
                $oSheet9->getRowDimension($ligne)->setRowHeight(35);
                $oSheet9->getStyle('A' . $ligne . ':R' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet9->getStyle('A' . $ligne . ':R' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet9->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($not, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(2, $ligne, $not['title_s'][0]);
                }
                if (array_key_exists("journalTitle_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(3, $ligne, $not['journalTitle_s']);
                }
                if (array_key_exists("producedDateY_i", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(4, $ligne, $not['producedDateY_i']);
                }
                $oSheet9->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet9->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet9->setCellValueByColumnAndRow(7, $ligne, getCoAuteursHceres($not));
                $oSheet9->setCellValueByColumnAndRow( 8,  $ligne, getPremierDernier($not,$equipes));
                $oSheet9->setCellValueByColumnAndRow(9, $ligne, getLangueHceres($not));
                if (array_key_exists("openAccess_bool", $not)) {
                    if ($not['openAccess_bool']) {
                        $oSheet9->setCellValueByColumnAndRow(10, $ligne, 'O');
                    } else {
                        $oSheet9->setCellValueByColumnAndRow(10, $ligne, 'N');
                    }
                }
                if (array_key_exists("volume_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(11, $ligne, $not['volume_s']);
                }
                if (array_key_exists("page_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(12, $ligne, $not['page_s']);
                }
                $oSheet9->setCellValueByColumnAndRow( 13,  $ligne, getCitationAutre($not));
                if (array_key_exists("abstract_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(14, $ligne, $not['abstract_s'][0]);
                }
                $oSheet9->setCellValueByColumnAndRow(15, $ligne, getKeyword($not));
                if (array_key_exists("doiId_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(16, $ligne, 'http://doi.org/' . $not['doiId_s']);
                    $oSheet9->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $not['doiId_s']);
                }
                if (array_key_exists("halId_s", $not)) {
                    $oSheet9->setCellValueByColumnAndRow(17, $ligne, 'https://hal.archives-ouvertes.fr/' . $not['halId_s']);
                    $oSheet9->getCellByColumnAndRow(17, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $not['halId_s']);
                }

                $oSheet9->getStyle('A'.$ligne.':Q'.$ligne)->getAlignment()->setVertical('top');
                $oSheet9->getStyle('D'.$ligne.':J'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet9->getStyle('G'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet9->getStyle('K'.$ligne.':L'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet9->getStyle('A' . $ligne . ':Q' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet9->getStyle('A' . $ligne . ':Q' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** TRAD *******************************//

        $oSheet10 = $oSpreadsheet->createSheet();

        $oSheet10->getTabColor()->setARGB('ffb1a0c7');
        $oSheet10->setTitle("Traduction");

        $oSheet10->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet10->getStyle('A1:N1')->applyFromArray($aStyleHeader);
        $oSheet10->getStyle('A1:N1')->getAlignment()->setWrapText(true);

        $oSheet10->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet10->mergeCells('A1:C1');

        $oSheet10->setCellValue('D1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet10->mergeCells('D1:G1');

        $oSheet10->setCellValue('H1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet10->mergeCells('H1:I1');

        $oSheet10->setCellValue('J1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet10->mergeCells('J1:N1');

        $oSheet10->getRowDimension(2)->setRowHeight(90);
        $oSheet10->getStyle('A2:N2')->applyFromArray($aStyleTitle);
        $oSheet10->setAutoFilter('A2:N2');
        $oSheet10->getStyle('A2:N2')->getAlignment()->setWrapText(true);
        $oSheet10->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet10->setCellValue('B2', 'Titre de la publication');    // colonne 1
        $oSheet10->setCellValue('C2', 'Année');    // colonne 3
        $oSheet10->setCellValue('D2', $richTextEquipe);    // colonne 6
        $oSheet10->setCellValue('E2', $richTextDoc);    // colonne 7
        $oSheet10->setCellValue('F2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet10->setCellValue('G2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet10->setCellValue('H2', 'Langue');    // colonne 10
        $oSheet10->setCellValue('I2', 'Article en "open access"');    // colonne 11
        $oSheet10->setCellValue('J2', 'Description');    // colonne 14
        $oSheet10->setCellValue('K2', 'Citation de l\'article');    // colonne 18
        $oSheet10->setCellValue('L2', 'Mots clés');    // colonne 4
        $oSheet10->setCellValue('M2', 'Lien DOI');    // colonne 19
        $oSheet10->setCellValue('N2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet10->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('F')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('J')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('K')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('M')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet10->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($trads)) {
            foreach ($trads as $trad) {
                $oSheet10->getRowDimension($ligne)->setRowHeight(35);
                $oSheet10->getStyle('A' . $ligne . ':N' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet10->getStyle('A' . $ligne . ':N' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet10->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($trad, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $trad)) {
                    $oSheet10->setCellValueByColumnAndRow(2, $ligne, $trad['title_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $trad)) {
                    $oSheet10->setCellValueByColumnAndRow(3, $ligne, $trad['producedDateY_i']);
                }
                $oSheet10->setCellValueByColumnAndRow(4, $ligne, '');
                $oSheet10->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet10->setCellValueByColumnAndRow(6, $ligne, getCoAuteursHceres($trad));
                $oSheet10->setCellValueByColumnAndRow( 7,  $ligne, getPremierDernier($trad,$equipes));
                $oSheet10->setCellValueByColumnAndRow(8, $ligne, getLangueHceres($trad));
                if (array_key_exists("openAccess_bool", $trad)) {
                    if ($trad['openAccess_bool']) {
                        $oSheet10->setCellValueByColumnAndRow(9, $ligne, 'O');
                    } else {
                        $oSheet10->setCellValueByColumnAndRow(9, $ligne, 'N');
                    }
                }
                if (array_key_exists("description_s", $trad)) {
                    $oSheet10->setCellValueByColumnAndRow(10, $ligne, $trad['description_s']);
                }
                $oSheet10->setCellValueByColumnAndRow( 11,  $ligne, getCitationAutre($trad));
                $oSheet10->setCellValueByColumnAndRow(12, $ligne, getKeyword($trad));
                if (array_key_exists("doiId_s", $trad)) {
                    $oSheet10->setCellValueByColumnAndRow(13, $ligne, 'http://doi.org/' . $trad['doiId_s']);
                    $oSheet10->getCellByColumnAndRow(13, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $trad['doiId_s']);
                }
                if (array_key_exists("halId_s", $trad)) {
                    $oSheet10->setCellValueByColumnAndRow(14, $ligne, 'https://hal.archives-ouvertes.fr/' . $trad['halId_s']);
                    $oSheet10->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $trad['halId_s']);
                }

                $oSheet10->getStyle('A'.$ligne.':R'.$ligne)->getAlignment()->setVertical('top');
                $oSheet10->getStyle('C'.$ligne.':I'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet10->getStyle('F'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet10->getStyle('G'.$ligne)->getAlignment()->setHorizontal('center');

                if ($ligne % 2 != 0) {
                    $oSheet10->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet10->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** BREVET *****************************//

        $oSheet11 = $oSpreadsheet->createSheet();

        $oSheet11->getTabColor()->setARGB('ffb1a0c7');
        $oSheet11->setTitle("Brevet");

        $oSheet11->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet11->getStyle('A1:N1')->applyFromArray($aStyleHeader);
        $oSheet11->getStyle('A1:N1')->getAlignment()->setWrapText(true);

        $oSheet11->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet11->mergeCells('A1:E1');

        $oSheet11->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet11->mergeCells('F1:I1');

        $oSheet11->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet11->mergeCells('J1:K1');

        $oSheet11->setCellValue('L1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet11->mergeCells('L1:N1');

        $oSheet11->getRowDimension(2)->setRowHeight(90);
        $oSheet11->getStyle('A2:N2')->applyFromArray($aStyleTitle);
        $oSheet11->setAutoFilter('A2:N2');
        $oSheet11->getStyle('A2:N2')->getAlignment()->setWrapText(true);
        $oSheet11->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet11->setCellValue('B2', 'Titre du brevet');    // colonne 1
        $oSheet11->setCellValue('C2', 'N°');    // colonne 16
        $oSheet11->setCellValue('D2', 'Pays');    // colonne 2
        $oSheet11->setCellValue('E2', 'Année');    // colonne 3
        $oSheet11->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet11->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet11->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet11->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet11->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet11->setCellValue('K2', 'Article en "open access"');    // colonne 11
        $oSheet11->setCellValue('L2', 'Citation de l\'article');    // colonne 18
        $oSheet11->setCellValue('M2', 'Lien DOI');    // colonne 19
        $oSheet11->setCellValue('N2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet11->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet11->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet11->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet11->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet11->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet11->getColumnDimension('M')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet11->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($brevets)) {
            foreach ($brevets as $brevet) {
                $oSheet11->getRowDimension($ligne)->setRowHeight(35);
                $oSheet11->getStyle('A' . $ligne . ':N' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet11->getStyle('A' . $ligne . ':N' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet11->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($brevet, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $brevet)) {
                    $oSheet11->setCellValueByColumnAndRow(2, $ligne, $brevet['title_s'][0]);
                }
                if (array_key_exists("number_s", $brevet)) {
                    $oSheet11->setCellValueByColumnAndRow(3, $ligne, $brevet['number_s'][0]);
                }
                if (array_key_exists("country_s", $brevet)) {
                    $oSheet11->setCellValueByColumnAndRow(4, $ligne, codeToCountry($brevet['country_s']));
                }
                if (array_key_exists("producedDateY_i", $brevet)) {
                    $oSheet11->setCellValueByColumnAndRow(5, $ligne, $brevet['producedDateY_i']);
                }
                $oSheet11->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet11->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet11->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($brevet));
                $oSheet11->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($brevet,$equipes));
                $oSheet11->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($brevet));
                if (array_key_exists("openAccess_bool", $brevet)) {
                    if ($brevet['openAccess_bool']) {
                        $oSheet11->setCellValueByColumnAndRow(11, $ligne, 'O');
                    } else {
                        $oSheet11->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                $oSheet11->setCellValueByColumnAndRow( 12,  $ligne, getCitationBrevet($brevet));
                if (array_key_exists("doiId_s", $brevet)) {
                    $oSheet11->setCellValueByColumnAndRow(13, $ligne, 'http://doi.org/' . $brevet['doiId_s']);
                    $oSheet11->getCellByColumnAndRow(13, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $brevet['doiId_s']);
                }
                if (array_key_exists("halId_s", $brevet)) {
                    $oSheet11->setCellValueByColumnAndRow(14, $ligne, 'https://hal.archives-ouvertes.fr/' . $brevet['halId_s']);
                    $oSheet11->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $brevet['halId_s']);
                }

                $oSheet11->getStyle('A'.$ligne.':N'.$ligne)->getAlignment()->setVertical('top');
                $oSheet11->getStyle('E'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet11->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet11->getStyle('C'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet11->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet11->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** AUTRE PUBLICATION ******************//

        $oSheet12 = $oSpreadsheet->createSheet();

        $oSheet12->getTabColor()->setARGB('ffb1a0c7');
        $oSheet12->setTitle("Autre publication");

        $oSheet12->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet12->getStyle('A1:M1')->applyFromArray($aStyleHeader);
        $oSheet12->getStyle('A1:M1')->getAlignment()->setWrapText(true);

        $oSheet12->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet12->mergeCells('A1:C1');

        $oSheet12->setCellValue('D1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet12->mergeCells('D1:G1');

        $oSheet12->setCellValue('H1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet12->mergeCells('H1:I1');

        $oSheet12->setCellValue('J1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet12->mergeCells('J1:M1');

        $oSheet12->getRowDimension(2)->setRowHeight(90);
        $oSheet12->getStyle('A2:M2')->applyFromArray($aStyleTitle);
        $oSheet12->setAutoFilter('A2:M2');
        $oSheet12->getStyle('A2:M2')->getAlignment()->setWrapText(true);
        $oSheet12->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet12->setCellValue('B2', 'Titre de la publication');    // colonne 1
        $oSheet12->setCellValue('C2', 'Année');    // colonne 3
        $oSheet12->setCellValue('D2', $richTextEquipe);    // colonne 6
        $oSheet12->setCellValue('E2', $richTextDoc);    // colonne 7
        $oSheet12->setCellValue('F2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet12->setCellValue('G2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet12->setCellValue('H2', 'Langue');    // colonne 10
        $oSheet12->setCellValue('I2', 'Article en "open access"');    // colonne 11
        $oSheet12->setCellValue('J2', 'Description');    // colonne 14
        $oSheet12->setCellValue('K2', 'Citation de l\'article');    // colonne 18
        $oSheet12->setCellValue('L2', 'Lien DOI');    // colonne 19
        $oSheet12->setCellValue('M2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet12->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('F')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('K')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('J')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('L')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet12->getColumnDimension('M')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($others)) {
            foreach ($others as $other) {
                $oSheet12->getRowDimension($ligne)->setRowHeight(35);
                $oSheet12->getStyle('A' . $ligne . ':M' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet12->getStyle('A' . $ligne . ':M' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet12->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($other, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $other)) {
                    $oSheet12->setCellValueByColumnAndRow(2, $ligne, $other['title_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $other)) {
                    $oSheet12->setCellValueByColumnAndRow(3, $ligne, $other['producedDateY_i']);
                }
                $oSheet12->setCellValueByColumnAndRow(4, $ligne, '');
                $oSheet12->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet12->setCellValueByColumnAndRow(6, $ligne, getCoAuteursHceres($other));
                $oSheet12->setCellValueByColumnAndRow( 7,  $ligne, getPremierDernier($other,$equipes));
                $oSheet12->setCellValueByColumnAndRow(8, $ligne, getLangueHceres($other));
                if (array_key_exists("openAccess_bool", $other)) {
                    if ($other['openAccess_bool']) {
                        $oSheet12->setCellValueByColumnAndRow(9, $ligne, 'O');
                    } else {
                        $oSheet12->setCellValueByColumnAndRow(9, $ligne, 'N');
                    }
                }
                if (array_key_exists("description_s", $other)) {
                    $oSheet12->setCellValueByColumnAndRow(10, $ligne, $other['description_s']);
                }
                $oSheet12->setCellValueByColumnAndRow( 11,  $ligne, getCitationAutre($other));
                if (array_key_exists("doiId_s", $other)) {
                    $oSheet12->setCellValueByColumnAndRow(12, $ligne, 'http://doi.org/' . $other['doiId_s']);
                    $oSheet12->getCellByColumnAndRow(12, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $other['doiId_s']);
                }
                if (array_key_exists("halId_s", $other)) {
                    $oSheet12->setCellValueByColumnAndRow(13, $ligne, 'https://hal.archives-ouvertes.fr/' . $other['halId_s']);
                    $oSheet12->getCellByColumnAndRow(13, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $other['halId_s']);
                }

                $oSheet12->getStyle('A'.$ligne.':M'.$ligne)->getAlignment()->setVertical('top');
                $oSheet12->getStyle('C'.$ligne.':I'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet12->getStyle('F'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet12->getStyle('A' . $ligne . ':M' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet12->getStyle('A' . $ligne . ':M' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** UNDEFINED **************************//

        $oSheet13 = $oSpreadsheet->createSheet();
        $oSheet13->getTabColor()->setARGB('ff92cddc');
        $oSheet13->setTitle("Preprint, Working Paper");

        $oSheet13->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet13->getStyle('A1:N1')->applyFromArray($aStyleHeader);
        $oSheet13->getStyle('A1:N1')->getAlignment()->setWrapText(true);

        $oSheet13->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet13->mergeCells('A1:E1');

        $oSheet13->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet13->mergeCells('F1:I1');

        $oSheet13->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet13->mergeCells('J1:K1');

        $oSheet13->setCellValue('L1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet13->mergeCells('L1:N1');

        $oSheet13->getRowDimension(2)->setRowHeight(90);
        $oSheet13->getStyle('A2:N2')->applyFromArray($aStyleTitle);
        $oSheet13->setAutoFilter('A2:N2');
        $oSheet13->getStyle('A2:N2')->getAlignment()->setWrapText(true);
        $oSheet13->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet13->setCellValue('B2', 'Titre de l\'article');    // colonne 1
        $oSheet13->setCellValue('C2', 'Nom de la revue');    // colonne 2
        $oSheet13->setCellValue('D2', 'Titre du volume ou de la collection');    // colonne 3
        $oSheet13->setCellValue('E2', 'Année');    // colonne 4
        $oSheet13->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet13->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet13->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet13->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet13->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet13->setCellValue('K2', 'Article en "open access"');    // colonne 11
        $oSheet13->setCellValue('L2', 'Citation de l\'article');    // colonne 18
        $oSheet13->setCellValue('M2', 'Lien DOI');    // colonne 19
        $oSheet13->setCellValue('N2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet13->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('M')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet13->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($prepublis)) {
            foreach ($prepublis as $prepubli) {
                $oSheet13->getRowDimension($ligne)->setRowHeight(35);
                $oSheet13->getStyle('A' . $ligne . ':N' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet13->getStyle('A' . $ligne . ':N' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet13->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($prepubli, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $prepubli)) {
                    $oSheet13->setCellValueByColumnAndRow(2, $ligne, $prepubli['title_s'][0]);
                }
                if (array_key_exists("journalTitle_s", $prepubli)) {
                    $oSheet13->setCellValueByColumnAndRow(3, $ligne, $prepubli['journalTitle_s']);
                }
                if (array_key_exists("serie_s", $prepubli)) {
                    $oSheet13->setCellValueByColumnAndRow(4, $ligne, $prepubli['serie_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $prepubli)) {
                    $oSheet13->setCellValueByColumnAndRow(5, $ligne, $prepubli['producedDateY_i']);
                }
                $oSheet13->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet13->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet13->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($prepubli));
                $oSheet13->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($prepubli,$equipes));
                $oSheet13->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($prepubli));
                if (array_key_exists("openAccess_bool", $prepubli)) {
                    if ($prepubli['openAccess_bool']) {
                        $oSheet13->setCellValueByColumnAndRow(11, $ligne, 'O');
                    } else {
                        $oSheet13->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                $oSheet13->setCellValueByColumnAndRow( 12,  $ligne, getCitationPrepubli($prepubli));
                if (array_key_exists("doiId_s", $prepubli)) {
                    $oSheet13->setCellValueByColumnAndRow(13, $ligne, 'http://doi.org/' . $prepubli['doiId_s']);
                    $oSheet13->getCellByColumnAndRow(13, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $prepubli['doiId_s']);
                }
                if (array_key_exists("halId_s", $prepubli)) {
                    $oSheet13->setCellValueByColumnAndRow(14, $ligne, 'https://hal.archives-ouvertes.fr/' . $prepubli['halId_s']);
                    $oSheet13->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $prepubli['halId_s']);
                }

                $oSheet13->getStyle('A'.$ligne.':N'.$ligne)->getAlignment()->setVertical('top');
                $oSheet13->getStyle('E'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet13->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet13->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet13->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** RAPPORT ****************************//

        $oSheet14 = $oSpreadsheet->createSheet();

        $oSheet14->getTabColor()->setARGB('ff92cddc');
        $oSheet14->setTitle("Rapport");

        $oSheet14->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet14->getStyle('A1:P1')->applyFromArray($aStyleHeader);
        $oSheet14->getStyle('A1:P1')->getAlignment()->setWrapText(true);

        $oSheet14->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet14->mergeCells('A1:D1');

        $oSheet14->setCellValue('E1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet14->mergeCells('E1:H1');

        $oSheet14->setCellValue('I1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet14->mergeCells('I1:J1');

        $oSheet14->setCellValue('K1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet14->mergeCells('K1:P1');

        $oSheet14->getRowDimension(2)->setRowHeight(90);
        $oSheet14->getStyle('A2:P2')->applyFromArray($aStyleTitle);
        $oSheet14->setAutoFilter('A2:P2');
        $oSheet14->getStyle('A2:P2')->getAlignment()->setWrapText(true);
        $oSheet14->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet14->setCellValue('B2', 'Titre du rapport');    // colonne 1
        $oSheet14->setCellValue('C2', 'Institution');    // colonne 2
        $oSheet14->setCellValue('D2', 'Année');    // colonne 3
        $oSheet14->setCellValue('E2', $richTextEquipe);    // colonne 6
        $oSheet14->setCellValue('F2', $richTextDoc);    // colonne 7
        $oSheet14->setCellValue('G2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet14->setCellValue('H2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet14->setCellValue('I2', 'Langue');    // colonne 10
        $oSheet14->setCellValue('J2', 'Article en "open access"');    // colonne 11
        $oSheet14->setCellValue('K2', 'N° du volume');    // colonne 15
        $oSheet14->setCellValue('L2', 'N°');    // colonne 16
        $oSheet14->setCellValue('M2', 'Pages');    // colonne 17
        $oSheet14->setCellValue('N2', 'Citation de l\'article');    // colonne 18
        $oSheet14->setCellValue('O2', 'Lien DOI');    // colonne 19
        $oSheet14->setCellValue('P2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet14->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('G')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet14->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($rapports)) {
            foreach ($rapports as $rapport) {
                $oSheet14->getRowDimension($ligne)->setRowHeight(35);
                $oSheet14->getStyle('A' . $ligne . ':P' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet14->getStyle('A' . $ligne . ':P' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet14->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($rapport, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(2, $ligne, $rapport['title_s'][0]);
                }
                if (array_key_exists("authorityInstitution_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(3, $ligne, $rapport['authorityInstitution_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(4, $ligne, $rapport['producedDateY_i']);
                }
                $oSheet14->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet14->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet14->setCellValueByColumnAndRow(7, $ligne, getCoAuteursHceres($rapport));
                $oSheet14->setCellValueByColumnAndRow( 8,  $ligne, getPremierDernier($rapport,$equipes));
                $oSheet14->setCellValueByColumnAndRow(9, $ligne, getLangueHceres($rapport));
                if (array_key_exists("openAccess_bool", $rapport)) {
                    if ($rapport['openAccess_bool']) {
                        $oSheet14->setCellValueByColumnAndRow(10, $ligne, 'O');
                    } else {
                        $oSheet14->setCellValueByColumnAndRow(10, $ligne, 'N');
                    }
                }
                if (array_key_exists("volume_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(11, $ligne, $rapport['volume_s']);
                }
                if (array_key_exists("number_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(12, $ligne, $rapport['number_s'][0]);
                }
                if (array_key_exists("page_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(13, $ligne, $rapport['page_s']);
                }
                $oSheet14->setCellValueByColumnAndRow( 14,  $ligne, getCitationRapport($rapport));
                if (array_key_exists("doiId_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(15, $ligne, 'http://doi.org/' . $rapport['doiId_s']);
                    $oSheet14->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $rapport['doiId_s']);
                }
                if (array_key_exists("halId_s", $rapport)) {
                    $oSheet14->setCellValueByColumnAndRow(16, $ligne, 'https://hal.archives-ouvertes.fr/' . $rapport['halId_s']);
                    $oSheet14->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $rapport['halId_s']);
                }

                $oSheet14->getStyle('A'.$ligne.':P'.$ligne)->getAlignment()->setVertical('top');
                $oSheet14->getStyle('D'.$ligne.':J'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet14->getStyle('G'.$ligne)->getAlignment()->setHorizontal('left');
                $oSheet14->getStyle('K'.$ligne.':M'.$ligne)->getAlignment()->setHorizontal('right');

                if ($ligne % 2 != 0) {
                    $oSheet14->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet14->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** THESE ******************************//

        $oSheet15 = $oSpreadsheet->createSheet();

        $oSheet15->getTabColor()->setARGB('fffabf8f');
        $oSheet15->setTitle("These");

        $oSheet15->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet15->getStyle('A1:N1')->applyFromArray($aStyleHeader);
        $oSheet15->getStyle('A1:N1')->getAlignment()->setWrapText(true);

        $oSheet15->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet15->mergeCells('A1:E1');

        $oSheet15->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet15->mergeCells('F1:I1');

        $oSheet15->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet15->mergeCells('J1:K1');

        $oSheet15->setCellValue('L1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet15->mergeCells('L1:N1');

        $oSheet15->getRowDimension(2)->setRowHeight(90);
        $oSheet15->getStyle('A2:N2')->applyFromArray($aStyleTitle);
        $oSheet15->setAutoFilter('A2:N2');
        $oSheet15->getStyle('A2:N2')->getAlignment()->setWrapText(true);
        $oSheet15->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet15->setCellValue('B2', 'Titre de la thèse');    // colonne 1
        $oSheet15->setCellValue('C2', 'Etablissement qui délivre le cours');    // colonne 3
        $oSheet15->setCellValue('D2', 'Année');    // colonne 3
        $oSheet15->setCellValue('E2', 'Date de soutenance');    // colonne 4
        $oSheet15->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet15->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet15->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet15->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet15->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet15->setCellValue('K2', 'Article en "open access"');    // colonne 11
        $oSheet15->setCellValue('L2', 'Citation de l\'article');    // colonne 18
        $oSheet15->setCellValue('M2', 'Lien DOI');    // colonne 19
        $oSheet15->setCellValue('N2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet15->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('M')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet15->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($theses)) {
            foreach ($theses as $these) {
                $oSheet15->getRowDimension($ligne)->setRowHeight(35);
                $oSheet15->getStyle('A' . $ligne . ':N' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet15->getStyle('A' . $ligne . ':N' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet15->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($these, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $these)) {
                    $oSheet15->setCellValueByColumnAndRow(2, $ligne, $these['title_s'][0]);
                }
                if (array_key_exists("authorityInstitution_s", $these)) {
                    $oSheet15->setCellValueByColumnAndRow(3, $ligne, $these['authorityInstitution_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $these)) {
                    $oSheet15->setCellValueByColumnAndRow(4, $ligne, $these['producedDateY_i']);
                }
                if (array_key_exists("producedDate_s", $these)) {
                    $oSheet15->setCellValueByColumnAndRow(5, $ligne, $these['producedDate_s']);
                }
                $oSheet15->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet15->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet15->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($these));
                $oSheet15->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($these,$equipes));
                $oSheet15->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($these));
                if (array_key_exists("openAccess_bool", $these)) {
                    if ($these['openAccess_bool']) {
                        $oSheet15->setCellValueByColumnAndRow(11, $ligne, 'O');
                    } else {
                        $oSheet15->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                $oSheet15->setCellValueByColumnAndRow( 12,  $ligne, getCitationTheseHdrMemoire($these));
                if (array_key_exists("doiId_s", $these)) {
                    $oSheet15->setCellValueByColumnAndRow(13, $ligne, 'http://doi.org/' . $these['doiId_s']);
                    $oSheet15->getCellByColumnAndRow(13, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $these['doiId_s']);
                }
                if (array_key_exists("halId_s", $these)) {
                    $oSheet15->setCellValueByColumnAndRow(14, $ligne, 'https://hal.archives-ouvertes.fr/' . $these['halId_s']);
                    $oSheet15->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $these['halId_s']);
                }

                $oSheet15->getStyle('A'.$ligne.':N'.$ligne)->getAlignment()->setVertical('top');
                $oSheet15->getStyle('D'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet15->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet15->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet15->getStyle('A' . $ligne . ':N' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** HDR ********************************//

        $oSheet16 = $oSpreadsheet->createSheet();

        $oSheet16->getTabColor()->setARGB('fffabf8f');
        $oSheet16->setTitle("HDR");

        $oSheet16->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet16->getStyle('A1:P1')->applyFromArray($aStyleHeader);
        $oSheet16->getStyle('A1:P1')->getAlignment()->setWrapText(true);

        $oSheet16->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet16->mergeCells('A1:E1');

        $oSheet16->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet16->mergeCells('F1:I1');

        $oSheet16->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet16->mergeCells('J1:K1');

        $oSheet16->setCellValue('L1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet16->mergeCells('L1:P1');

        $oSheet16->getRowDimension(2)->setRowHeight(90);
        $oSheet16->getStyle('A2:P2')->applyFromArray($aStyleTitle);
        $oSheet16->setAutoFilter('A2:P2');
        $oSheet16->getStyle('A2:P2')->getAlignment()->setWrapText(true);
        $oSheet16->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet16->setCellValue('B2', 'Titre HDR');    // colonne 1
        $oSheet16->setCellValue('C2', 'Etablissement qui délivre le cours');    // colonne 3
        $oSheet16->setCellValue('D2', 'Année');    // colonne 3
        $oSheet16->setCellValue('E2', 'Date de soutenance');    // colonne 4
        $oSheet16->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet16->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet16->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet16->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet16->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet16->setCellValue('K2', 'Article en "open access"');    // colonne 11
        $oSheet16->setCellValue('L2', 'Citation de l\'article');    // colonne 18
        $oSheet16->setCellValue('M2', 'Résumé');    // colonne 4
        $oSheet16->setCellValue('N2', 'Mots clés');    // colonne 4
        $oSheet16->setCellValue('O2', 'Lien DOI');    // colonne 19
        $oSheet16->setCellValue('P2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet16->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('M')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet16->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($hdrs)) {
            foreach ($hdrs as $hdr) {
                $oSheet16->getRowDimension($ligne)->setRowHeight(35);
                $oSheet16->getStyle('A' . $ligne . ':P' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet16->getStyle('A' . $ligne . ':P' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet16->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($hdr, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(2, $ligne, $hdr['title_s'][0]);
                }
                if (array_key_exists("authorityInstitution_s", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(3, $ligne, $hdr['authorityInstitution_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(4, $ligne, $hdr['producedDateY_i']);
                }
                if (array_key_exists("producedDate_s", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(5, $ligne, $hdr['producedDate_s']);
                }
                $oSheet16->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet16->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet16->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($hdr));
                $oSheet16->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($hdr,$equipes));
                $oSheet16->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($hdr));
                if (array_key_exists("openAccess_bool", $hdr)) {
                    if ($hdr['openAccess_bool']) {
                        $oSheet16->setCellValueByColumnAndRow(11, $ligne, 'O');
                    } else {
                        $oSheet16->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                $oSheet16->setCellValueByColumnAndRow( 12,  $ligne, getCitationTheseHdrMemoire($hdr));
                if (array_key_exists("abstract_s", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(13, $ligne, $hdr['abstract_s'][0]);
                }
                $oSheet16->setCellValueByColumnAndRow(14, $ligne, getKeyword($hdr));
                if (array_key_exists("doiId_s", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(15, $ligne, 'http://doi.org/' . $hdr['doiId_s']);
                    $oSheet16->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $hdr['doiId_s']);
                }
                if (array_key_exists("halId_s", $hdr)) {
                    $oSheet16->setCellValueByColumnAndRow(16, $ligne, 'https://hal.archives-ouvertes.fr/' . $hdr['halId_s']);
                    $oSheet16->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $hdr['halId_s']);
                }

                $oSheet16->getStyle('A'.$ligne.':P'.$ligne)->getAlignment()->setVertical('top');
                $oSheet16->getStyle('D'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet16->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet16->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet16->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** COURS ******************************//

        $oSheet17 = $oSpreadsheet->createSheet();

        $oSheet17->getTabColor()->setARGB('fffabf8f');
        $oSheet17->setTitle("Cours");

        $oSheet17->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet17->getStyle('A1:R1')->applyFromArray($aStyleHeader);
        $oSheet17->getStyle('A1:R1')->getAlignment()->setWrapText(true);

        $oSheet17->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet17->mergeCells('A1:G1');

        $oSheet17->setCellValue('H1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet17->mergeCells('H1:K1');

        $oSheet17->setCellValue('L1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet17->mergeCells('L1:M1');

        $oSheet17->setCellValue('N1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet17->mergeCells('N1:R1');

        $oSheet17->getRowDimension(2)->setRowHeight(90);
        $oSheet17->getStyle('A2:R2')->applyFromArray($aStyleTitle);
        $oSheet17->setAutoFilter('A2:R2');
        $oSheet17->getStyle('A2:R2')->getAlignment()->setWrapText(true);
        $oSheet17->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet17->setCellValue('B2', 'Titre du cours');    // colonne 1
        $oSheet17->setCellValue('C2', 'Nom du cours');    // colonne 1
        $oSheet17->setCellValue('D2', 'Niveau du cours');    // colonne 1
        $oSheet17->setCellValue('E2', 'Etablissement qui délivre le cours');    // colonne 3
        $oSheet17->setCellValue('F2', 'Pays');    // colonne 3
        $oSheet17->setCellValue('G2', 'Année');    // colonne 3
        $oSheet17->setCellValue('H2', $richTextEquipe);    // colonne 6
        $oSheet17->setCellValue('I2', $richTextDoc);    // colonne 7
        $oSheet17->setCellValue('J2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet17->setCellValue('K2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet17->setCellValue('L2', 'Langue');    // colonne 10
        $oSheet17->setCellValue('M2', 'Article en "open access"');    // colonne 11
        $oSheet17->setCellValue('N2', 'Citation de l\'article');    // colonne 18
        $oSheet17->setCellValue('O2', 'Résumé');    // colonne 4
        $oSheet17->setCellValue('P2', 'Mots clés');    // colonne 4
        $oSheet17->setCellValue('Q2', 'Lien DOI');    // colonne 19
        $oSheet17->setCellValue('R2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet17->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('E')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('J')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('O')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('P')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('Q')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet17->getColumnDimension('R')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($cours)) {
            foreach ($cours as $cour) {
                $oSheet17->getRowDimension($ligne)->setRowHeight(35);
                $oSheet17->getStyle('A' . $ligne . ':R' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet17->getStyle('A' . $ligne . ':R' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet17->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($cour, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(2, $ligne, $cour['title_s'][0]);
                }
                if (array_key_exists("lectureName_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(3, $ligne, $cour['lectureName_s']);
                }
                if (array_key_exists("lectureType_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(4, $ligne, getLectureType($cour['lectureType_s']));
                }
                if (array_key_exists("authorityInstitution_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(5, $ligne, $cour['authorityInstitution_s'][0]);
                }
                if (array_key_exists("country_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(6, $ligne, codeToCountry($cour['country_s']));
                }
                if (array_key_exists("producedDateY_i", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(7, $ligne, $cour['producedDateY_i']);
                }
                $oSheet17->setCellValueByColumnAndRow(8, $ligne, '');
                $oSheet17->setCellValueByColumnAndRow(9, $ligne, '');
                $oSheet17->setCellValueByColumnAndRow(10, $ligne, getCoAuteursHceres($cour));
                $oSheet17->setCellValueByColumnAndRow( 11,  $ligne, getPremierDernier($cour,$equipes));
                $oSheet17->setCellValueByColumnAndRow(12, $ligne, getLangueHceres($cour));
                if (array_key_exists("openAccess_bool", $cour)) {
                    if ($cour['openAccess_bool']) {
                        $oSheet17->setCellValueByColumnAndRow(13, $ligne, 'O');
                    } else {
                        $oSheet17->setCellValueByColumnAndRow(13, $ligne, 'N');
                    }
                }
                $oSheet17->setCellValueByColumnAndRow( 14,  $ligne, getCitationCours($cour));
                if (array_key_exists("abstract_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(15, $ligne, $cour['abstract_s'][0]);
                }
                $oSheet17->setCellValueByColumnAndRow(16, $ligne, getKeyword($cour));
                if (array_key_exists("doiId_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(17, $ligne, 'http://doi.org/' . $cour['doiId_s']);
                    $oSheet17->getCellByColumnAndRow(17, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $cour['doiId_s']);
                }
                if (array_key_exists("halId_s", $cour)) {
                    $oSheet17->setCellValueByColumnAndRow(18, $ligne, 'https://hal.archives-ouvertes.fr/' . $cour['halId_s']);
                    $oSheet17->getCellByColumnAndRow(18, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $cour['halId_s']);
                }

                $oSheet17->getStyle('A'.$ligne.':R'.$ligne)->getAlignment()->setVertical('top');
                $oSheet17->getStyle('G'.$ligne.':M'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet17->getStyle('J'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet17->getStyle('A' . $ligne . ':R' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet17->getStyle('A' . $ligne . ':R' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** MEDIA ******************************//

        if(empty($images)){
            $images = [];
        }

        if(empty($videos)){
            $videos = [];
        }

        if(empty($sons)){
            $sons = [];
        }

        if(empty($cartes)){
            $cartes = [];
        }

        $medias = array_merge($images, $videos, $sons, $cartes);

        $oSheet18 = $oSpreadsheet->createSheet();

        $oSheet18->getTabColor()->setARGB('ffda9694');
        $oSheet18->setTitle("Media");

        $oSheet18->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet18->getStyle('A1:P1')->applyFromArray($aStyleHeader);
        $oSheet18->getStyle('A1:P1')->getAlignment()->setWrapText(true);

        $oSheet18->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet18->mergeCells('A1:E1');

        $oSheet18->setCellValue('F1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet18->mergeCells('F1:I1');

        $oSheet18->setCellValue('J1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet18->mergeCells('J1:K1');

        $oSheet18->setCellValue('L1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet18->mergeCells('L1:P1');

        $oSheet18->getRowDimension(2)->setRowHeight(90);
        $oSheet18->getStyle('A2:P2')->applyFromArray($aStyleTitle);
        $oSheet18->setAutoFilter('A2:P2');
        $oSheet18->getStyle('A2:P2')->getAlignment()->setWrapText(true);
        $oSheet18->setCellValue('A2', 'Type de document');    // colonne 0
        $oSheet18->setCellValue('B2', 'Auteurs');    // colonne 0
        $oSheet18->setCellValue('C2', 'Titre du document');    // colonne 1
        $oSheet18->setCellValue('D2', 'Année');    // colonne 3
        $oSheet18->setCellValue('E2', 'Date de création');    // colonne 4
        $oSheet18->setCellValue('F2', $richTextEquipe);    // colonne 6
        $oSheet18->setCellValue('G2', $richTextDoc);    // colonne 7
        $oSheet18->setCellValue('H2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet18->setCellValue('I2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet18->setCellValue('J2', 'Langue');    // colonne 10
        $oSheet18->setCellValue('K2', 'Article en "open access"');    // colonne 11
        $oSheet18->setCellValue('L2', 'Citation de l\'article');    // colonne 18
        $oSheet18->setCellValue('M2', 'Résumé');    // colonne 4
        $oSheet18->setCellValue('N2', 'Mots clés');    // colonne 4
        $oSheet18->setCellValue('O2', 'Lien DOI');    // colonne 19
        $oSheet18->setCellValue('P2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet18->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('C')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('H')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('M')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('N')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet18->getColumnDimension('P')->setWidth(50+self::WIDTH_ADAPTATOR);

        if (!empty($medias)) {
            foreach ($medias as $media) {
                $oSheet18->getRowDimension($ligne)->setRowHeight(35);
                $oSheet18->getStyle('A' . $ligne . ':P' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet18->getStyle('A' . $ligne . ':P' . $ligne)->getAlignment()->setWrapText(true);
                if (array_key_exists("title_s", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(1, $ligne, getTypeMedia($media));
                }
                $oSheet18->setCellValueByColumnAndRow(2, $ligne, getAuteursSoulign($media, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(3, $ligne, $media['title_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(4, $ligne, $media['producedDateY_i']);
                }
                if (array_key_exists("producedDate_s", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(5, $ligne, $media['producedDate_s']);
                }
                $oSheet18->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet18->setCellValueByColumnAndRow(7, $ligne, '');
                $oSheet18->setCellValueByColumnAndRow(8, $ligne, getCoAuteursHceres($media));
                $oSheet18->setCellValueByColumnAndRow( 9,  $ligne, getPremierDernier($media,$equipes));
                $oSheet18->setCellValueByColumnAndRow(10, $ligne, getLangueHceres($media));
                if (array_key_exists("openAccess_bool", $media)) {
                    if ($media['openAccess_bool']) {
                        $oSheet18->setCellValueByColumnAndRow(11, $ligne, 'O');
                    } else {
                        $oSheet18->setCellValueByColumnAndRow(11, $ligne, 'N');
                    }
                }
                $oSheet18->setCellValueByColumnAndRow( 12,  $ligne, getCitationMedia($media));
                if (array_key_exists("abstract_s", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(13, $ligne, $media['abstract_s'][0]);
                }
                $oSheet18->setCellValueByColumnAndRow(14, $ligne, getKeyword($media));
                if (array_key_exists("doiId_s", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(15, $ligne, 'http://doi.org/' . $media['doiId_s']);
                    $oSheet18->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $media['doiId_s']);
                }
                if (array_key_exists("halId_s", $media)) {
                    $oSheet18->setCellValueByColumnAndRow(16, $ligne, 'https://hal.archives-ouvertes.fr/' . $media['halId_s']);
                    $oSheet18->getCellByColumnAndRow(16, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $media['halId_s']);
                }

                $oSheet18->getStyle('A'.$ligne.':P'.$ligne)->getAlignment()->setVertical('top');
                $oSheet18->getStyle('D'.$ligne.':K'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet18->getStyle('H'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet18->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet18->getStyle('A' . $ligne . ':P' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        ////////////////////////////////////////////////////////////////
        // ********************** LOGICIEL ***************************//

        $oSheet22 = $oSpreadsheet->createSheet();

        $oSheet22->getTabColor()->setARGB('ffda9694');
        $oSheet22->setTitle("Logiciel");

        $oSheet22->getRowDimension(1)->setRowHeight(16.5, "px");
        $oSheet22->getStyle('A1:O1')->applyFromArray($aStyleHeader);
        $oSheet22->getStyle('A1:O1')->getAlignment()->setWrapText(true);

        $oSheet22->setCellValue('A1', 'Description de l\'article');    // colonne 0
        $oSheet22->mergeCells('A1:D1');

        $oSheet22->setCellValue('E1', 'Caractéristiques des auteurs');    // colonne 0
        $oSheet22->mergeCells('E1:H1');

        $oSheet22->setCellValue('I1', 'Caractéristiques de l\'article');    // colonne 0
        $oSheet22->mergeCells('I1:J1');

        $oSheet22->setCellValue('K1', 'Eléments complémentaires de l\'article');    // colonne 0
        $oSheet22->mergeCells('K1:O1');

        $oSheet22->getRowDimension(2)->setRowHeight(90);
        $oSheet22->getStyle('A2:O2')->applyFromArray($aStyleTitle);
        $oSheet22->setAutoFilter('A2:O2');
        $oSheet22->getStyle('A2:O2')->getAlignment()->setWrapText(true);
        $oSheet22->setCellValue('A2', 'Auteurs');    // colonne 0
        $oSheet22->setCellValue('B2', 'Nom du logiciel');    // colonne 1
        $oSheet22->setCellValue('C2', 'Année');    // colonne 3
        $oSheet22->setCellValue('D2', 'Date de production/écriture');    // colonne 4
        $oSheet22->setCellValue('E2', $richTextEquipe);    // colonne 6
        $oSheet22->setCellValue('F2', $richTextDoc);    // colonne 7
        $oSheet22->setCellValue('G2', 'Affiliation institutionnelle des co-auteurs');    // colonne 8
        $oSheet22->setCellValue('H2', 'Publication en premier, dernier ou auteur de correspondance');    // colonne 9
        $oSheet22->setCellValue('I2', 'Langue');    // colonne 10
        $oSheet22->setCellValue('J2', 'Article en "open access"');    // colonne 11
        $oSheet22->setCellValue('K2', 'Citation de l\'article');    // colonne 18
        $oSheet22->setCellValue('L2', 'Résumé');    // colonne 4
        $oSheet22->setCellValue('M2', 'Mots clés');    // colonne 4
        $oSheet22->setCellValue('N2', 'Lien DOI');    // colonne 19
        $oSheet22->setCellValue('O2', 'Lien HAL');    // colonne 20

        $ligne = 3;  // on remplit l'onglet avec les données à partir de la ligne n°2

        $oSheet22->getDefaultColumnDimension()->setWidth(20+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('A')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('B')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('G')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('K')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('L')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('M')->setWidth(80+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('N')->setWidth(50+self::WIDTH_ADAPTATOR);
        $oSheet22->getColumnDimension('O')->setWidth(50+self::WIDTH_ADAPTATOR);


        if (!empty($logs)) {
            foreach ($logs as $log) {
                $oSheet22->getRowDimension($ligne)->setRowHeight(35);
                $oSheet22->getStyle('A' . $ligne . ':O' . $ligne)->applyFromArray($aStyleNormal);
                $oSheet22->getStyle('A' . $ligne . ':O' . $ligne)->getAlignment()->setWrapText(true);
                $oSheet22->setCellValueByColumnAndRow(1, $ligne, getAuteursSoulign($log, $soulignAuteur, $equipes));
                if (array_key_exists("title_s", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(2, $ligne, $log['title_s'][0]);
                }
                if (array_key_exists("producedDateY_i", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(3, $ligne, $log['producedDateY_i']);
                }
                if (array_key_exists("producedDate_s", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(4, $ligne, $log['producedDate_s']);
                }
                $oSheet22->setCellValueByColumnAndRow(5, $ligne, '');
                $oSheet22->setCellValueByColumnAndRow(6, $ligne, '');
                $oSheet22->setCellValueByColumnAndRow(7, $ligne, getCoAuteursHceres($log));
                $oSheet22->setCellValueByColumnAndRow( 8,  $ligne, getPremierDernier($log,$equipes));
                $oSheet22->setCellValueByColumnAndRow(9, $ligne, getLangueHceres($log));
                if (array_key_exists("openAccess_bool", $log)) {
                    if ($log['openAccess_bool']) {
                        $oSheet22->setCellValueByColumnAndRow(10, $ligne, 'O');
                    } else {
                        $oSheet22->setCellValueByColumnAndRow(10, $ligne, 'N');
                    }
                }
                $oSheet22->setCellValueByColumnAndRow( 11,  $ligne, getCitationLogiciel($log));
                if (array_key_exists("abstract_s", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(12, $ligne, $log['abstract_s'][0]);
                }
                if (array_key_exists("keyword_s", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(13, $ligne, $log['keyword_s'][0]);
                }
                if (array_key_exists("doiId_s", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(14, $ligne, 'http://doi.org/' . $log['doiId_s']);
                    $oSheet22->getCellByColumnAndRow(14, $ligne)->getHyperlink()->setUrl('http://doi.org/' . $log['doiId_s']);
                }
                if (array_key_exists("halId_s", $log)) {
                    $oSheet22->setCellValueByColumnAndRow(15, $ligne, 'https://hal.archives-ouvertes.fr/' . $log['halId_s']);
                    $oSheet22->getCellByColumnAndRow(15, $ligne)->getHyperlink()->setUrl('https://hal.archives-ouvertes.fr/' . $log['halId_s']);
                }

                $oSheet22->getStyle('A'.$ligne.':O'.$ligne)->getAlignment()->setVertical('top');
                $oSheet22->getStyle('C'.$ligne.':J'.$ligne)->getAlignment()->setHorizontal('center');
                $oSheet22->getStyle('G'.$ligne)->getAlignment()->setHorizontal('left');

                if ($ligne % 2 != 0) {
                    $oSheet22->getStyle('A' . $ligne . ':O' . $ligne)->getFill()->getStartColor()->setARGB('ffd9d9d9');
                }else{
                    $oSheet22->getStyle('A' . $ligne . ':O' . $ligne)->getFill()->setFillType(Fill::FILL_NONE);
                }

                $ligne++;
                $nbNoticeTraite++;
            }
        }

        if(!empty($articles)){
            $countArt = count($articles);

            $coverSheet->setCellValue('A14','✓');
            $coverSheet->getStyle('A14')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A14')->getFont()->setBold(true);
        }else{
            $countArt = 0;

            $coverSheet->setCellValue('A14','✘');
            $coverSheet->getStyle('A14')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($comms)){
            $countComm = count($comms);

            $coverSheet->setCellValue('A15','✓');
            $coverSheet->getStyle('A15')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A15')->getFont()->setBold(true);
        }else{
            $countComm = 0;

            $coverSheet->setCellValue('A15','✘');
            $coverSheet->getStyle('A15')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($posters)){
            $countPost = count($posters);

            $coverSheet->setCellValue('A16','✓');
            $coverSheet->getStyle('A16')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A16')->getFont()->setBold(true);
        }else{
            $countPost = 0;

            $coverSheet->setCellValue('A16','✘');
            $coverSheet->getStyle('A16')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($procs)){
            $countProc = count($procs);

            $coverSheet->setCellValue('A17','✓');
            $coverSheet->getStyle('A17')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A17')->getFont()->setBold(true);
        }else{
            $countProc = 0;

            $coverSheet->setCellValue('A17','✘');
            $coverSheet->getStyle('A17')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($issues)){
            $countIss = count($issues);

            $coverSheet->setCellValue('A18','✓');
            $coverSheet->getStyle('A18')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A18')->getFont()->setBold(true);
        }else{
            $countIss = 0;

            $coverSheet->setCellValue('A18','✘');
            $coverSheet->getStyle('A18')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($ouvrages)){
            $countOuv = count($ouvrages);

            $coverSheet->setCellValue('A19','✓');
            $coverSheet->getStyle('A19')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A19')->getFont()->setBold(true);
        }else{
            $countOuv = 0;

            $coverSheet->setCellValue('A19','✘');
            $coverSheet->getStyle('A19')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($chapitres)){
            $countCouv = count($chapitres);

            $coverSheet->setCellValue('A20','✓');
            $coverSheet->getStyle('A20')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A20')->getFont()->setBold(true);
        }else{
            $countCouv = 0;

            $coverSheet->setCellValue('A20','✘');
            $coverSheet->getStyle('A20')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($blogs)){
            $countBlog = count($blogs);

            $coverSheet->setCellValue('A21','✓');
            $coverSheet->getStyle('A21')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A21')->getFont()->setBold(true);
        }else{
            $countBlog = 0;

            $coverSheet->setCellValue('A21','✘');
            $coverSheet->getStyle('A21')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($nots)){
            $countNot = count($nots);

            $coverSheet->setCellValue('A22','✓');
            $coverSheet->getStyle('A22')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A22')->getFont()->setBold(true);
        }else{
            $countNot = 0;

            $coverSheet->setCellValue('A22','✘');
            $coverSheet->getStyle('A22')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($trads)){
            $countTrad = count($trads);

            $coverSheet->setCellValue('A23','✓');
            $coverSheet->getStyle('A23')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A23')->getFont()->setBold(true);
        }else{
            $countTrad = 0;

            $coverSheet->setCellValue('A23','✘');
            $coverSheet->getStyle('A23')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($brevets)){
            $countBrev = count($brevets);

            $coverSheet->setCellValue('A24','✓');
            $coverSheet->getStyle('A24')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A24')->getFont()->setBold(true);
        }else{
            $countBrev = 0;

            $coverSheet->setCellValue('A24','✘');
            $coverSheet->getStyle('A24')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($others)){
            $countOth = count($others);

            $coverSheet->setCellValue('A25','✓');
            $coverSheet->getStyle('A25')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('A25')->getFont()->setBold(true);
        }else{
            $countOth = 0;

            $coverSheet->setCellValue('A25','✘');
            $coverSheet->getStyle('A25')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($prepublis)){
            $countPre = count($prepublis);

            $coverSheet->setCellValue('D14','✓');
            $coverSheet->getStyle('D14')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('D14')->getFont()->setBold(true);
        }else{
            $countPre = 0;

            $coverSheet->setCellValue('D14','✘');
            $coverSheet->getStyle('D14')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($rapports)){
            $countRap = count($rapports);

            $coverSheet->setCellValue('D15','✓');
            $coverSheet->getStyle('D15')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('D15')->getFont()->setBold(true);
        }else{
            $countRap = 0;

            $coverSheet->setCellValue('D15','✘');
            $coverSheet->getStyle('D15')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($theses)){
            $countThes = count($theses);

            $coverSheet->setCellValue('G14','✓');
            $coverSheet->getStyle('G14')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('G14')->getFont()->setBold(true);
        }else{
            $countThes = 0;

            $coverSheet->setCellValue('G14','✘');
            $coverSheet->getStyle('G14')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($hdrs)){
            $countHdr = count($hdrs);

            $coverSheet->setCellValue('G15','✓');
            $coverSheet->getStyle('G15')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('G15')->getFont()->setBold(true);
        }else{
            $countHdr = 0;

            $coverSheet->setCellValue('G15','✘');
            $coverSheet->getStyle('G15')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($cours)){
            $countCou = count($cours);

            $coverSheet->setCellValue('G16','✓');
            $coverSheet->getStyle('G16')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('G16')->getFont()->setBold(true);
        }else{
            $countCou = 0;

            $coverSheet->setCellValue('G16','✘');
            $coverSheet->getStyle('G16')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($medias)){
            $countMed = count($medias);

            $coverSheet->setCellValue('J14','✓');
            $coverSheet->getStyle('J14')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('J14')->getFont()->setBold(true);
        }else{
            $countMed = 0;

            $coverSheet->setCellValue('J14','✘');
            $coverSheet->getStyle('J14')->getFont()->getColor()->setARGB('ffff0000');
        }

        if(!empty($logs)){
            $countLog = count($logs);

            $coverSheet->setCellValue('J15','✓');
            $coverSheet->getStyle('J15')->getFont()->getColor()->setARGB('ff00b050');
            $coverSheet->getStyle('J15')->getFont()->setBold(true);
        }else{
            $countLog = 0;

            $coverSheet->setCellValue('J15','✘');
            $coverSheet->getStyle('J15')->getFont()->getColor()->setARGB('ffff0000');
        }

        $sheetCount = $oSpreadsheet->getSheetCount();
        for ($i = 2; $i < $sheetCount; $i++) {
            $sheet = $oSpreadsheet->getSheet($i);
            $sheet->setSelectedCell('A3');
            $sheet->freezePane('A3');
        }

        $aCategories = array(
            array("Article dans une revue",$countArt),
            array("Communication dans un congrès",$countComm),
            array("Poster",$countPost),
            array("Proceedings/Recueil des communications",$countProc),
            array("No spécial de revue/special issue",$countIss),
            array("Ouvrage (y compris édition critique et traduction)",$countOuv),
            array("Chapitres dʹouvrage",$countCouv),
            array("Article de blog scientifique",$countBlog),
            array("Notice dʹencyclopédie ou dictionnaire",$countNot),
            array("Traduction",$countTrad),
            array("Brevets",$countBrev),
            array("Autre publication",$countOth),
            array("Pré-publication, Document de travail",$countPre),
            array("Rapport",$countRap),
            array("Thèse",$countThes),
            array("HDR",$countHdr),
            array("Cours",$countCou),
            array("Media",$countMed),
            array("Logiciel",$countLog),
        );

        $sMsg = "<table class='striped'>";
        $sCodeStats = "";
        $iCompteur = 0;
        foreach($aCategories as $aLigne){
            if ($iCompteur==0) $sMsg .= "<tr>";
            $sMsg .= "<td class='titreCat'>" . $aLigne[0] . "</td><td>" . $aLigne[1] . "</td>";
            $iCompteur = ($iCompteur==2) ?  0 : ($iCompteur+1);
            if ($iCompteur==0) $sMsg .= "</tr>";
            $sCodeStats .= (($sCodeStats!='') ? ',' : '') . sprintf("{ y: %d, name: '%s', exploded: true }", $aLigne[1], $aLigne[0]);
        }
        $sMsg .= "</table>";


        $oWriter = new Xlsx($oSpreadsheet);

        //Désactivation du code VBA dans le fichier - report en 2023
        $macroCode = file_get_contents("../src/Helper/vba/vbaProject.bin");
        $oSpreadsheet->setMacrosCode($macroCode);

        // enregistrement du fichier xlsx
        $oWriter->save($myXls);

        $urlFile = $myXls;

        $retourne[0] = $urlHAL;
        $retourne[1] = $sMsg;
        $retourne[2] = $urlFile;
        $retourne[3] = $nbNoticeTraite;
        $retourne[4] = "[" . $sCodeStats . "]";

        if ($debug) {
            foreach ($articles as $article)
                fwrite($h, print_r($article, true));
        }
        return json_encode($retourne);   // on retourne en json les infos à afficher dans la page appelante

    }
}



