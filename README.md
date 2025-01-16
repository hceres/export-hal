# Exporter la production scientifique de votre laboratoire depuis l'archive ouverte HAL
 
## Objectif et usages
Ce code source motorise une application web qui permet à un laboratoire de recherche d'exporter sa production scientifique référencée dans HAL (https://hal.science/) dans un fichier Excel facile à lire.
 
En quelques clics, cet export contribue à simplifier la collecte d'informations sur les unités de recherche dans le cadre de leur évaluation par le Hcéres (Haut Conseil de l'évaluation de la recherche et de l'enseignement supérieur, www.hceres.fr). Lors de la campagne d'évaluation 2024-2025, l’application web a été utilisée par près de 80% des laboratoires évalués. Présentée au club utilisateur de l'archive ouverte HAL (https://www.casuhal.org/), cette application est utilisée toute l'année, en dehors du cadre de l'évaluation du Hcéres (environ 70 exports par jour).
 
## Licence ouverte
Cette application repose sur un code source ouvert et initialement développé par l'INRAE, dans le cadre de ses propres procédures interne d'évaluation, à partir de l’outil  ExtrHAL (Université de Rennes).
Le présent code source est distribué sous licence GNU GPL
 
## Fonctionnalités
L'application propose un formulaire simple (cf. https://monevaluation.hceres.fr/hal) où le laboratoire de recherche peut saisir les informations suivantes :
- Année de début et de fin, pour limiter l'export aux productions scientifiques publiées entre ces deux années (incluses). Dans le cadre de l'évaluation des entités de recherche, l'export est paramétré sur les 5 années qui sont examinées dans la campagne d’évaluation en cours.
- L'identification du laboratoire recherche, soit par son sigle, soit par le code de sa collection HAL, ou par son identifiant de référence HAL AuréHAL (https://aurehal.archives-ouvertes.fr/structure/index ).
 
Le bouton "Rechercher" affiche le résultat de la recherche dans la page Web sous la forme d'un tableau et d'un diagramme, indiquant par catégorie combien de productions scientifiques ont été trouvées dans l’archive HAL par appel de l’API de recherche HAL (API HAL : https://api.archives-ouvertes.fr/docs/search) .
 
Le bouton "Télécharger ma liste de productions" permet d'exporter cette liste dans un format Excel, où la production est répartie dans les onglets par type de support de publication. La nomenclature des types de support de publication utilisée est celle fournie par HAL et dont la liste est la suivante :
- Article dans une revue
- Communication dans un congrès
- Poster
- Proceedings/Recueil des communications
- No spécial de revue/special issue
- Ouvrage (y compris édition critique et traduction)
- Chapitres dʹouvrage
- Article de blog scientifique
- Notice dʹencyclopédie ou dictionnaire
- Traduction
- Brevets
- Autre publication
- Pré-publication, Document de travail
- Rapport
- Thèse
- HDR
- Cours
- Media
- Logiciel
 
Le fichier Excel comporte également un onglet "Repérer équipes et doctorants". Dès lors que les listes nominatives des personnels et des doctorants de l'unité sont remplies dans cet onglet, actionner le bouton "Repérer les équipes et doctorants de mon UR parmi les auteurs des productions" permet d'alimenter les colonnes F et G de chaque onglet avec les équipes et les doctorants de chaque publication travaillant au sein du laboratoire de recherche.
 
## Architecture technique
Ce code source s'appuie sur la version 5.4 du framework Symfony. Il est compatible avec PHP 7 ou 8. 
Composer est utilisé afin de gérer toutes les dépendances du projet. Ces dernières sont installables avec la commande "composer install".
Le fichier ".env" nécessaire au fonctionnement peut être instancié à partir du fichier .env.dist. Les valeurs des variables d'environnement sont données à titre d'exemple. Le fichier ".env" doit se trouver à la racine du projet, au même endroit que le .env.dist.
 
### Captcha de l'État
Le formulaire peut être protégé par un captcha. Le code s’appuie sur le captcha de l'État (https://api.gouv.fr/les-api/api-captchetat). Si vous ne pouvez pas disposer de compte sur CaptchEtat, vous pouvez le désactiver dans le fichier ".env".
 
### Suivi d'activité
Le suivi de l'usage de l'application est réalisé via Matomo. Vous pouvez paramétrer votre propre instance en renseignant les variables d'environnement PROD_URL_COLLECTE et PROD_ID.
