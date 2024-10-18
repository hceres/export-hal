<?php

/* Fonctions pour l'export HCERES
calcul des citations pour chaque onglet du fichier excel
*/

// fichier comportant la fonction pays => language
use PhpOffice\PhpSpreadsheet\RichText\RichText;

@include_once('fct_lang.php');

//Suppresion des accents
function wdRemoveAccents($str, $charset='utf-8')
{
	$str = htmlentities($str, ENT_NOQUOTES, $charset);
	$str = preg_replace('#&([A-za-z])(?:acute|cedil|circ|grave|orn|ring|slash|th|tilde|uml);#', '\1', $str);
	$str = preg_replace('#&([A-za-z]{2})(?:lig);#', '\1', $str); // pour les ligatures e.g. '&oelig;'
	return preg_replace('#&[^;]+;#', '', $str); // supprime les autres caractères
}

function getAuteursHceres($notice)
{
	$i = 0;
	$ret = '';
	foreach ($notice['authFirstName_s'] as $prenom) {
		if (isPrenomCompose($notice['authFullName_s'][$i])) {
            $ret .= strtoupper(wdRemoveAccents($notice['authLastName_s'][$i])) .' '. getInitialCompose(wdRemoveAccents($notice['authFullName_s'][$i])) .', ';
        }else {
            $ret .= strtoupper(wdRemoveAccents($notice['authLastName_s'][$i])) .' '. wdRemoveAccents(prenomCompInit($prenom)) .', ';
        }
		$i++;
	}
	$ret = substr($ret, 0, -2);

	return $ret;
}

// $nomComplet = authFullName_s
// $nom        = authLastName_s
// $prenom     = authFirstName_s
function getNomAuteur($nom, $prenom, $nomComplet)
{
	if (isPrenomCompose($nomComplet)) {
		return strtoupper(wdRemoveAccents($nom)) .' '. getInitialCompose(wdRemoveAccents($nomComplet));
	}else {
		return strtoupper(wdRemoveAccents($nom)) .' '. wdRemoveAccents(prenomCompInit($prenom));
	}
}

// $labo = $equipes = 1002209~574738~74637~100299 liste de docid structure à partir du champ Informations laboratoires
// $soulignement : booléen pour activer le soulignement des auteurs
function getAuteursSoulign($notice, $soulignement, $labo)
{
	if ($soulignement == 0) {
        return getAuteursHceres($notice);
    }else {

		$objRichText = new RichText();
		$nbAuteur = count($notice['authFirstName_s']);
		$tabATester = explode('~', $labo);

		$aTabAuteurASouligner = array();
		// on construit d'abord le tableau des docid des auteurs à souligner
		foreach ($notice['authIdHasPrimaryStructure_fs'] as $auteurPS) {	// pour chaque auteur rattaché à une structure primaire
			$tabAuteur = explode('_JoinSep_', $auteurPS); // on sépare l'auteur de sa structure => $tabAuteur[0] contient la partie auteur ex: 1643348_FacetSep_Patrick Baldet
			// $tabAuteur[1] contient la partie structure ex: 566980_FacetSep_Écosystèmes forestiers
			$tabAuteurDocid  = explode('_FacetSep_', $tabAuteur[0]);
			$tabAuteurStruct = explode('_FacetSep_', $tabAuteur[1]);
			$idStructAuteur = $tabAuteurStruct[0];   // contient le DOCID de la structure de l'auteur

			$aTmpId = array();

			if (in_array($idStructAuteur, $tabATester)) {
				//On crée un tableau dans lequel on reconstitue le code de l'auteur id_FacetSep_Nom Prenom
				$aTmpId = explode('-', $tabAuteurDocid[0]);
			}

			if (!empty($aTmpId)) {
				$aTabAuteurASouligner[] = $aTmpId[1] . '_FacetSep_' . $tabAuteurDocid[1];
			}
		}

		$i=0;
		foreach ($notice['authIdFullName_fs'] as $auteur) {    // pour chaque auteur  12596751_FacetSep_N. Morellet
			$nomAuteur = getNomAuteur($notice['authLastName_s'][$i], $notice['authFirstName_s'][$i], $notice['authFullName_s'][$i]);

			if (in_array($auteur, $aTabAuteurASouligner)) { // le docid de l'auteur est dans le tableau des auteurs à souligner
				$objUnderlined = $objRichText->createTextRun($nomAuteur);
				$objUnderlined->getFont()->setBold(true);
			} else {
				$objRichText->createText($nomAuteur);
			}

			if ($i < $nbAuteur-1) {
				$objRichText->createText(', ');
			}
			$i++;
		}

		return $objRichText;
	}
}


function getLangueHceres($notice)
{
	if (isset($notice['language_s'])) {
		return codeToLanguage($notice['language_s'][0]);
	} else {
		if (isset($notice['country_s'])) {
			return codeToCountry($notice['country_s']);
		}else {
			return '';
		}
	}
}

// Retourne la liste des affiliations qu'il faut comparer à celles des auteurs
// cette liste est obtenue à partir de ce que l'utilisateur a saisi dans le champ EquipeLabo
function getAffiliations($equipe)
{
	$tabEquipeLabo = explode('~', $equipe); //%7E <=> ~
	$lstStruct = '';
	
	foreach ($tabEquipeLabo as $equipeLabo) {

		$req = 'https'.'://api.archives-ouvertes.fr/ref/structure/?q=(name_t:'.$equipeLabo.'%20OR%20acronym_t:'.$equipeLabo.')%20AND%20valid_s:(VALID%20OR%20OLD)%20AND%20country_s:%22fr%22&fl=docid';
		// Exemple avec ECOBIO retourne 1 valeur : docid=928 https://api.archives-ouvertes.fr/ref/structure/?q=(name_t:ECOBIO%20OR%20acronym_t:ECOBIO)%20AND%20valid_s:(VALID%20OR%20OLD)%20AND%20country_s:%22fr%22&fl=docid
		
		$resultat = file_get_contents($req);
		$res = json_decode($resultat);
		foreach ($res->response->docs as $valeur) {
			$lstStruct .= $valeur->docid.'~';
		}
	}

	return $lstStruct;
}

function getIndexAuteur($notice, $docid)
{
	$i=0;
	$index = 0;
	foreach ($notice['authIdFullName_fs'] as $auteur) {
		$docidReel = explode('-', $docid)[1];

		if ($docidReel == explode('_', $auteur)[0]) {
			$index = $i;
		}

		$i++;
	}
	return $index; // on retourne l'index de l'auteur dans ce tableau qui servira à retrouver son role dans authQuality_s
}

// Parametres : une notice et la liste des structures à tester $equipes = getAffiliations(...)
// si Identifiant structure (($equipeLabo) saisi est "1002209~AGIR" alors $equipes = 1002209~574738~74637~100299~560019~1002209~552012~225605~531394
// Ex de valeur pour authIdHasPrimaryStructure_fs ["1643348_FacetSep_Patrick Baldet_JoinSep_566980_FacetSep_Écosystèmes forestiers","11996480_FacetSep_Fabienne Colas_JoinSep_1021679_FacetSep_Direction de la recherche forestière"]
// dans chaque case du tableau: un auteur, pour chaque auteur on a DOCID auteur FacetSep Prénom Nom _JoinSep_ DOCID Structure FacetSep Nom structure
// Ex de valeur pour authQuality_s ["aut","crp"]
// Ex authLastName_s":["Baldet","Colas"]
function getPremierDernier($notice, $lstEquipe)
{
	$retour = 'N';

	// si rien a comparé
	if ($lstEquipe == '') {
		return $retour;
	}

	$i=0;
	$nbAuteur = count($notice['authQuality_s']);
	$tabATester = explode('~', $lstEquipe);
	
	foreach ($notice['authIdHasPrimaryStructure_fs'] as $auteur) {	// pour chaque auteur affilié
		$tabAuteur = explode('_JoinSep_', $auteur); // on sépare l'auteur de sa structure $tabAuteur[0] contient la partie auteur ex: 1643348_FacetSep_Patrick Baldet
		// $tabAuteur[1] contient la partie structure ex: 566980_FacetSep_Écosystèmes forestiers
		$tabAuteurStruct = explode('_FacetSep_', $tabAuteur[1]);
        $tabAuteurDocid  = explode('_FacetSep_', $tabAuteur[0]);
		$idStructAuteur = $tabAuteurStruct[0];   // contient le DOCID de la structure de l'auteur
		$docidAuteur = $tabAuteurDocid[0];   // contient le DOCID de l'auteur affilié
		
		$indexAuteur = getIndexAuteur($notice, $docidAuteur);
		
		if ($indexAuteur == 0 && in_array($idStructAuteur, $tabATester)) { // on teste le premier auteur et son appartenance à la liste $lstEquipe
            $retour='O';
		}

		if ($indexAuteur == $nbAuteur-1 && in_array($idStructAuteur, $tabATester)) {  // on teste le dernier auteur et son appartenance à la liste $lstEquipe
			$retour='O';
		}
		
		if ($notice['authQuality_s'][$indexAuteur] == 'crp' && in_array($idStructAuteur, $tabATester)) {  // on teste l'auteur correspondant
            $retour='O';
		}

		if ($notice['authQuality_s'][$indexAuteur] == 'co_first_author' && in_array($idStructAuteur, $tabATester)) {  // on teste le premier co-auteur
            $retour='O';
		}

		if ($notice['authQuality_s'][$indexAuteur] == 'co_last_author' && in_array($idStructAuteur, $tabATester)) {  // on teste le dernier co-auteur
            $retour='O';
		}

		$i++;
	}

	return $retour;
}

function getAudience($notice)
{
	if (isset($notice['audience_s'])) {
		switch ($notice['audience_s']) {
			case '1' :
                $return =  'Non spécifié';
                break;
			case '2' :
                $return =  'Internationale';
                break;
			case '3' :
                $return =  'Nationale';
                break;
            default:
                $return =  '';
		}

        return $return;
	}
}

function getTypeMedia($notice)
{
	if (isset($notice['docType_s'])) {
		switch ($notice['docType_s']) {
			case 'IMG' :
                $return = 'Image';
                break;
			case 'VIDEO' :
                $return =  'Vidéo';
                break;
			case 'SON' :
                $return =  'Son';
                break;
			case 'MAP' :
                $return =  'Carte';
                break;
            default:
                $return =  '';
		}

        return $return;
	}
}

/*************************** CITATION  **************************************/
/****************************************************************************/

function getCitationArticle($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	$double = false;

	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ssTitre= $notice['subTitle_s'][0].'. ';
		$ret .= $ssTitre;
	}

	if (isset($notice['journalTitle_s'])) {
		$journal= $notice['journalTitle_s'].', ';
		$ret .= $journal;
	}

	if (isset($notice['volume_s'])) {
		$ret .= $notice['volume_s'];
	}else {
		// attention double espace si pas de volume
		$double=true;
	}
	// mettre une virgule apres le volume si pas d'issue_s
	if (isset($notice['issue_s']) && $double) {
		$ret .= '('.$notice['issue_s'][0].'), ';
	}

	if (isset($notice['issue_s']) && !$double) {
		$ret .= ' ('.$notice['issue_s'][0].'), ';
	}
	
	if (!isset($notice['issue_s'])) {
		$ret .= ', ';
	}
	
	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'].', ';
	}
	
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationOuvrage($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}
		
	if (isset($notice['publisher_s'][0])) {
		$ret .=  $notice['publisher_s'][0] . ', ';
	}
	
	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'] . ', ';
	}
	
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationChapitreOuvrage($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}
	
	if (isset($notice['bookTitle_s'])) {
		$ret .=  'In : '.$notice['bookTitle_s'] . '. ';
	}
		
	if (isset($notice['publisher_s'][0])) {
		$ret .=  $notice['publisher_s'][0] . ', ';
	}
	
	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'] . ', ';
	}
	
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationCommunication($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	if (isset($notice['title_s'])) {
		$ret .=  $notice['title_s'][0].'. ';
	}
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}
	
	if (isset($notice['conferenceTitle_s'])) {
		$ret .=  'Presented at '.$notice['conferenceTitle_s'].', ';
	}
	
	if (isset($notice['city_s'])) {
		$ret .=  $notice['city_s'].', ';
	}
	
	if (isset($notice['country_s'])) {
		$ret .=  codeToCountry($notice['country_s']).' ';
	}
	
	if (isset($notice['conferenceStartDate_s'])) {
		$ret .=  '('.$notice['conferenceStartDate_s'].'), ';
	}
	
	if (isset($notice['invitedCommunication_s']) && $notice['invitedCommunication_s'] == '1') {
		$ret .=  ' Conférence invitée. ';
	}
	
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
		
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationBrevet($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .= $notice['subTitle_s'][0].'. ';
	}
		
	if (isset($notice['number_s'])) {
		$ret .= '(Brevet n°: '.$notice['number_s'][0].'). ';
	}
		
	if (isset($notice['country_s'])) {
		$ret .=  codeToCountry($notice['country_s']).', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationRapport($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .= $notice['subTitle_s'][0].'. ';
	}
		
	if (isset($notice['authorityInstitution_s'])) {
		$ret .= '('.$notice['authorityInstitution_s'][0].'), ';
	}
	
	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'].', ';
	}
	
	if (isset($notice['number_s'])) {
		$ret .= $notice['number_s'][0].', ';
	}
		
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
		
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationPrepubli($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}
	
	if (isset($notice['serie_s'][0])) {
		$ret .=  $notice['serie_s'][0] . ', ';
	}
	
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
	
	if (isset($notice['arxivId_s'])) {
		$ret .= 'n° arxiv: '.$notice['arxivId_s'].', ';
	}
		
	if (isset($notice['biorxivId_s'])) {
		$ret .= 'n° biorxiv: '.$notice['biorxivId_s'][0].', ';
	}

	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationTheseHdrMemoire($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].' ';
	}
	
	if (isset($notice['authorityInstitution_s'][0])) {
		$ret .= '('.$notice['authorityInstitution_s'][0].'). ';
	}
	
	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'].', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationCours($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].' ';
	}
	
	if (isset($notice['lectureName_s']) || isset($notice['lectureType_s']) || isset($notice['authorityInstitution_s'])) {
		$cours = '('. $notice['lectureName_s'] .', '. getLectureType($notice['lectureType_s']) .', '. $notice['authorityInstitution_s'][0] .')';
		$cours = str_replace('(, ', '(', $cours);
		$cours = str_replace(', )', ')', $cours);
		$ret .= $cours.', ';
	}
	
	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'].', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationLogiciel($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}
	
	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}
	
	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}
	
	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationMedia($notice)
{
	$ret = getAuteursHceres($notice). ' ';
	
	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}

	$ret .= $notice['title_s'][0].'. ';
	
	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}

	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}

	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCitationAutre($notice)
{
	$ret = getAuteursHceres($notice). ' ';

	if (isset($notice['producedDate_s'])) {
		$ret .= '('.$notice['producedDate_s'].'). ';
	}

	$ret .= $notice['title_s'][0].'. ';

	if (isset($notice['subTitle_s'])) {
		$ret .=  $notice['subTitle_s'][0].'. ';
	}

	if (isset($notice['authorityInstitution_s'][0])) {
		$ret .= '('.$notice['authorityInstitution_s'][0].'). ';
	}

	if (isset($notice['country_s'])) {
		$ret .= codeToCountry($notice['country_s']).', ';
	}

	if (isset($notice['serie_s'][0])) {
		$ret .=  $notice['serie_s'][0] . ', ';
	}

	if (isset($notice['journalTitle_s'])) {
		$ret .= $notice['journalTitle_s'].', ';
	}

	if (isset($notice['bookTitle_s'])) {
		$ret .= 'In : ' . $notice['bookTitle_s'] . '. ';
	}

	if (isset($notice['lectureName_s']) || isset($notice['lectureType_s']) || isset($notice['authorityInstitution_s'])) {
		$cours = '('. $notice['lectureName_s'] .', '. getLectureType($notice['lectureType_s']) .', '. $notice['authorityInstitution_s'][0] .')';
		$cours = str_replace('(, ', '(', $cours);
		$cours = str_replace(', )', ')', $cours);
		$ret .= $cours.', ';
	}
		
	if (isset($notice['publisher_s'][0])) {
		$ret .=  $notice['publisher_s'][0] . ', ';
	}

	if (isset($notice['arxivId_s'])) {
		$ret .= 'n° arxiv: '.$notice['arxivId_s'].', ';
	}

	if (isset($notice['biorxivId_s'])) {
		$ret .= 'n° biorxiv: '.$notice['biorxivId_s'][0].', ';
	}

	if (isset($notice['page_s'])) {
		$ret .= $notice['page_s'].', ';
	}

	if (isset($notice['doiId_s'])) {
		$ret .= $notice['doiId_s'].', ';
	}

	$lien = changeUri($notice['halId_s']);
	$ret .= $lien;
	
	return $ret;
}

function getCoAuteursHceres($notice)
{

	$res = '';

	foreach ($notice['authIdHasPrimaryStructure_fs'] as $coauteur) {
		$values = preg_split("/[\n\r_]+/", $coauteur);
		$res .= $values[2] . ': ' . $values[6] . ' ;';
	}

	return substr($res, 0, -2);
}

function getKeyword($notice)
{
	$res = '';
	if (isset($notice['keyword_s'])) {
		foreach ($notice['keyword_s'] as $keyword) {
			$res .= $keyword . ', ';
		}
		$res = substr($res, 0, -2);
	}

	return $res;
}
