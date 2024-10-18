<?php

/************************************************************/
// Fonctions permettant l'affichage des 19 types de document
/************************************************************/

function changeUri($halId)
{
    // L’url d’origine : https://hal-univ-lyon1.archives-ouvertes.fr/hal-02016398
    // Devrait être : https://hal.archives-ouvertes.fr/hal-0201639
    return 'https://hal.archives-ouvertes.fr/'.$halId;
}

function codeToCountry($code)
{
    $code = strtoupper($code);
    $countryList = array(
	'EN' => 'Anglais',
    'AF' => 'Afghanistan',
    'ZA' => 'Afrique du Sud',
    'AL' => 'Albanie',
    'DZ' => 'Algérie',
    'DE' => 'Allemagne',
    'AD' => 'Andorre',
    'AO' => 'Angola',
    'AI' => 'Anguilla',
    'AQ' => 'Antarctique',
    'AG' => 'Antigua-et-Barbuda',
    'SA' => 'Arabie saoudite',
    'AR' => 'Argentine',
    'AM' => 'Arménie',
    'AW' => 'Aruba',
    'AU' => 'Australie',
    'AT' => 'Autriche',
    'AZ' => 'Azerbaïdjan',
    'BS' => 'Bahamas',
    'BH' => 'Bahreïn',
    'BD' => 'Bangladesh',
    'BB' => 'Barbade',
    'BE' => 'Belgique',
    'BZ' => 'Belize',
    'BJ' => 'Bénin',
    'BM' => 'Bermudes',
    'BT' => 'Bhoutan',
    'BY' => 'Biélorussie',
    'BO' => 'Bolivie',
    'BA' => 'Bosnie-Herzégovine',
    'BW' => 'Botswana',
    'BR' => 'Brésil',
    'BN' => 'Brunei',
    'BG' => 'Bulgarie',
    'BF' => 'Burkina Faso',
    'BI' => 'Burundi',
    'KH' => 'Cambodge',
    'CM' => 'Cameroun',
    'CA' => 'Canada',
    'CV' => 'Cap-Vert',
    'CL' => 'Chili',
    'CN' => 'Chine',
    'CY' => 'Chypre',
    'CO' => 'Colombie',
    'KM' => 'Comores',
    'CG' => 'Congo-Brazzaville',
    'CD' => 'Congo-Kinshasa',
    'KP' => 'Corée du Nord',
    'KR' => 'Corée du Sud',
    'CR' => 'Costa Rica',
    'CI' => 'Côte d’Ivoire',
    'HR' => 'Croatie',
    'CU' => 'Cuba',
    'CW' => 'Curaçao',
    'DK' => 'Danemark',
    'DJ' => 'Djibouti',
    'DM' => 'Dominique',
    'EG' => 'Égypte',
    'AE' => 'Émirats arabes unis',
    'EC' => 'Équateur',
    'ER' => 'Érythrée',
    'ES' => 'Espagne',
    'EE' => 'Estonie',
    'SZ' => 'Eswatini',
    'VA' => 'État de la Cité du Vatican',
    'US' => 'États-Unis',
    'ET' => 'Éthiopie',
    'FJ' => 'Fidji',
    'FI' => 'Finlande',
    'FR' => 'France',
    'GA' => 'Gabon',
    'GM' => 'Gambie',
    'GE' => 'Géorgie',
    'GS' => 'Géorgie du Sud-et-les Îles Sandwich du Sud',
    'GH' => 'Ghana',
    'GI' => 'Gibraltar',
    'GR' => 'Grèce',
    'GD' => 'Grenade',
    'GL' => 'Groenland',
    'GP' => 'Guadeloupe',
    'GU' => 'Guam',
    'GT' => 'Guatemala',
    'GG' => 'Guernesey',
    'GN' => 'Guinée',
    'GQ' => 'Guinée équatoriale',
    'GW' => 'Guinée-Bissau',
    'GY' => 'Guyana',
    'GF' => 'Guyane française',
    'HT' => 'Haïti',
    'HN' => 'Honduras',
    'HU' => 'Hongrie',
    'BV' => 'Île Bouvet',
    'CX' => 'Île Christmas',
    'IM' => 'Île de Man',
    'NF' => 'Île Norfolk',
    'AX' => 'Îles Åland',
    'KY' => 'Îles Caïmans',
    'CC' => 'Îles Cocos',
    'CK' => 'Îles Cook',
    'FO' => 'Îles Féroé',
    'HM' => 'Îles Heard-et-MacDonald',
    'FK' => 'Îles Malouines',
    'MP' => 'Îles Mariannes du Nord',
    'MH' => 'Îles Marshall',
    'UM' => 'Îles mineures éloignées des États-Unis',
    'PN' => 'Îles Pitcairn',
    'SB' => 'Îles Salomon',
    'TC' => 'Îles Turques-et-Caïques',
    'VG' => 'Îles Vierges britanniques',
    'VI' => 'Îles Vierges des États-Unis',
    'IN' => 'Inde',
    'ID' => 'Indonésie',
    'IQ' => 'Irak',
    'IR' => 'Iran',
    'IE' => 'Irlande',
    'IS' => 'Islande',
    'IL' => 'Israël',
    'IT' => 'Italie',
    'JM' => 'Jamaïque',
    'JP' => 'Japon',
    'JE' => 'Jersey',
    'JO' => 'Jordanie',
    'KZ' => 'Kazakhstan',
    'KE' => 'Kenya',
    'KG' => 'Kirghizstan',
    'KI' => 'Kiribati',
    'KW' => 'Koweït',
    'RE' => 'La Réunion',
    'LA' => 'Laos',
    'LS' => 'Lesotho',
    'LV' => 'Lettonie',
    'LB' => 'Liban',
    'LR' => 'Liberia',
    'LY' => 'Libye',
    'LI' => 'Liechtenstein',
    'LT' => 'Lituanie',
    'LU' => 'Luxembourg',
    'MK' => 'Macédoine du Nord',
    'MG' => 'Madagascar',
    'MY' => 'Malaisie',
    'MW' => 'Malawi',
    'MV' => 'Maldives',
    'ML' => 'Mali',
    'MT' => 'Malte',
    'MA' => 'Maroc',
    'MQ' => 'Martinique',
    'MU' => 'Maurice',
    'MR' => 'Mauritanie',
    'YT' => 'Mayotte',
    'MX' => 'Mexique',
    'FM' => 'Micronésie',
    'MD' => 'Moldavie',
    'MC' => 'Monaco',
    'MN' => 'Mongolie',
    'ME' => 'Monténégro',
    'MS' => 'Montserrat',
    'MZ' => 'Mozambique',
    'MM' => 'Myanmar (Birmanie)',
    'NA' => 'Namibie',
    'NR' => 'Nauru',
    'NP' => 'Népal',
    'NI' => 'Nicaragua',
    'NE' => 'Niger',
    'NG' => 'Nigeria',
    'NU' => 'Niue',
    'NO' => 'Norvège',
    'NC' => 'Nouvelle-Calédonie',
    'NZ' => 'Nouvelle-Zélande',
    'OM' => 'Oman',
    'UG' => 'Ouganda',
    'UZ' => 'Ouzbékistan',
    'PK' => 'Pakistan',
    'PW' => 'Palaos',
    'PA' => 'Panama',
    'PG' => 'Papouasie-Nouvelle-Guinée',
    'PY' => 'Paraguay',
    'NL' => 'Pays-Bas',
    'BQ' => 'Pays-Bas caribéens',
    'PE' => 'Pérou',
    'PH' => 'Philippines',
    'PL' => 'Pologne',
    'PF' => 'Polynésie française',
    'PR' => 'Porto Rico',
    'PT' => 'Portugal',
    'QA' => 'Qatar',
    'HK' => 'R.A.S. chinoise de Hong Kong',
    'MO' => 'R.A.S. chinoise de Macao',
    'CF' => 'République centrafricaine',
    'DO' => 'République dominicaine',
    'RO' => 'Roumanie',
    'GB' => 'Royaume-Uni',
    'RU' => 'Russie',
    'RW' => 'Rwanda',
    'EH' => 'Sahara occidental',
    'BL' => 'Saint-Barthélemy',
    'KN' => 'Saint-Christophe-et-Niévès',
    'SM' => 'Saint-Marin',
    'MF' => 'Saint-Martin',
    'SX' => 'Saint-Martin (partie néerlandaise)',
    'PM' => 'Saint-Pierre-et-Miquelon',
    'VC' => 'Saint-Vincent-et-les Grenadines',
    'SH' => 'Sainte-Hélène',
    'LC' => 'Sainte-Lucie',
    'SV' => 'Salvador',
    'WS' => 'Samoa',
    'AS' => 'Samoa américaines',
    'ST' => 'Sao Tomé-et-Principe',
    'SN' => 'Sénégal',
    'RS' => 'Serbie',
    'SC' => 'Seychelles',
    'SL' => 'Sierra Leone',
    'SG' => 'Singapour',
    'SK' => 'Slovaquie',
    'SI' => 'Slovénie',
    'SO' => 'Somalie',
    'SD' => 'Soudan',
    'SS' => 'Soudan du Sud',
    'LK' => 'Sri Lanka',
    'SE' => 'Suède',
    'CH' => 'Suisse',
    'SR' => 'Suriname',
    'SJ' => 'Svalbard et Jan Mayen',
    'SY' => 'Syrie',
    'TJ' => 'Tadjikistan',
    'TW' => 'Taïwan',
    'TZ' => 'Tanzanie',
    'TD' => 'Tchad',
    'CZ' => 'Tchéquie',
    'TF' => 'Terres australes françaises',
    'IO' => 'Territoire britannique de l’océan Indien',
    'PS' => 'Territoires palestiniens',
    'TH' => 'Thaïlande',
    'TL' => 'Timor oriental',
    'TG' => 'Togo',
    'TK' => 'Tokelau',
    'TO' => 'Tonga',
    'TT' => 'Trinité-et-Tobago',
    'TN' => 'Tunisie',
    'TM' => 'Turkménistan',
    'TR' => 'Turquie',
    'TV' => 'Tuvalu',
    'UA' => 'Ukraine',
    'UY' => 'Uruguay',
    'VU' => 'Vanuatu',
    'VE' => 'Venezuela',
    'VN' => 'Viêt Nam',
    'WF' => 'Wallis-et-Futuna',
    'YE' => 'Yémen',
    'ZM' => 'Zambie',
    'ZW' => 'Zimbabwe',
    );

    if (!$countryList[$code]) {
        return $code;
    }else {
        return $countryList[$code];
    }

}

// correspondance pour le champ lectureType_s
function getLectureType($code)
{
	switch ($code) {
		case 1 :
            $return = 'DEA';
            break;
		case 2 :
            $return = 'École thématique';
            break;
		case 7 :
            $return = '3ème cycle';
            break;
		case 10 :
            $return = 'École d\'ingénieur';
            break;
		case 11 :
            $return = 'Licence';
            break;
		case 12 :
            $return = 'Master';
            break;
		case 13 :
            $return = 'Doctorat';
            break;
		case 14 :
            $return = 'DEUG';
            break;
		case 15 :
            $return = 'Maîtrise';
            break;
		case 21 :
            $return = 'Licence/L1';
            break;
		case 22 :
            $return = 'Licence/L2';
            break;
		case 23 :
            $return = 'Licence/L3';
            break;
		case 31 :
            $return = 'Master/M1';
            break;
		case 32 :
            $return = 'Master/M2';
            break;
		case 40 :
            $return = 'Vulgarisation';
            break;
        default:
            $return = '';
	}
    
    return $return;
}

function mbUcwords($str)
{
  $str = mb_convert_case($str, MB_CASE_TITLE, 'UTF-8');
  return ($str);
}

function prenomCompInit($prenom)
{
    $prenom = str_replace('  ', ' ', $prenom);

    if (strpos(trim($prenom), '-') !== false) {//Le prénom comporte un tiret
        $postiret = mb_strpos(trim($prenom), '-', 0, 'UTF-8');

        if ($postiret != 1) {
            $prenomg = trim(mb_substr($prenom, 0, ($postiret-1), 'UTF-8'));
        }else {
            $prenomg = trim(mb_substr($prenom, 0, 1, 'UTF-8'));
        }

        $prenomd = trim(mb_substr($prenom, ($postiret+1), strlen($prenom), 'UTF-8'));
        $autg = mb_substr($prenomg, 0, 1, 'UTF-8');
        $autd = mb_substr($prenomd, 0, 1, 'UTF-8');
        $prenom = mbUcwords($autg).'.-'.mbUcwords($autd).'.';
    }else {
        if (strpos(trim($prenom), ' ') !== false) {//plusieurs prénoms
            $tabprenom = explode(' ', trim($prenom));
            $p = 0;
            $prenom = '';
            while (isset($tabprenom[$p])) {
                if ($p == 0) {
                    $prenom .= mbUcwords(mb_substr($tabprenom[$p], 0, 1, 'UTF-8')).'.';
                }else {
                    $prenom .= ' '.mbUcwords(mb_substr($tabprenom[$p], 0, 1, 'UTF-8')).'.';
                }
                $p++;
            }
        }else {
            $prenom = mbUcwords(mb_substr($prenom, 0, 1, 'UTF-8')).'.';
        }
    }
    return $prenom;
}

function isPrenomCompose($nom)
{
    // on veut savoir si le prénom est composé de deux intiales, ex : J.F. Picard
	$tab = explode(' ', $nom);

   	return substr_count($tab[0], '.') == 2 && strlen($tab[0]) == 4;
}

function getInitialCompose($nom)
{
	// si le nom est de la forme J.F. Bernard alors on retourne Bernard J.F.
	$tab = explode('. ', $nom);
	return ($tab[0].'.');
}

