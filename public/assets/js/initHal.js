var preLoader = '';

var changeHeaderOn=50;

function EnteteAttachHandler(){
	// console.log($("#ContentHolder").height());

	var cbpAnimatedHeader = (function() {

		var docElem = document.documentElement,
			header = document.querySelector('#TopContentHolder'),
			didScroll = false;
			//changeHeaderOn = 240;

		function init() {
			window.addEventListener( 'scroll', function( event ) {
				if( !didScroll ) {
					didScroll = true;
					scrollPage();
				}
			}, false );
			if(scrollY()>0) scrollPage();
		}

		function scrollPage() {
			var sy = scrollY();
			if ( sy >= changeHeaderOn ) {
				if (!$(header).hasClass("NavPinned z-depth-1")){
					$(header).switchClass("", "NavPinned z-depth-1");
				}
			} else {
				$(header).switchClass("NavPinned z-depth-1", "");
			}
			didScroll = false;
		}

		function scrollY() {
			return window.pageYOffset || docElem.scrollTop;
		}

		init();
	})();
	$(".breadcrumClick, .menuItem", "#TopContentHolder").off("click").on("click", function(){
		var sAttrV = $(this).attr("relv");
		var sAttr = $(this).attr("rel");
		if ((typeof sAttr !== typeof undefined) && (sAttr !== false)) {
			if ((typeof sAttrV !== typeof undefined) && (sAttrV !== false)) {
				$.globalEval($(this).attr("rel")+"('"+sAttrV+"');");
			}else{
				$.globalEval($(this).attr("rel")+"();");
			}
		}
	});
}

function efface(formeBool){
	$('#resultats').html(' ');
	$('#error-content').html(' ');
	$('#details').html(' ');
	$('#error-section').hide();

}

function getIdsHal() {
	var formatedIds = "";
	$('.chip').each(function(index){
		var separator = index === 0 ? "" : "-";
		formatedIds+=separator+$(this).attr('hal-value');
	})
	return formatedIds;
}

function setError(content) {
	$('#error-section').show();
	$('#error-content').html(content);
}

function lance_recherche(){

	//Anticipation utilisateur ne clique pas sur "ajouter" (identifiant AuréHal)
	$( "#add-id-hal" ).trigger('click');

	// on commence par effacer les précédents resultats
	efface(false);
	
	var idcoll = "";
	var equipe = "";
	var idshal = "";
	var dated  = $('#datedeb').val().trim();
	var datef  = $('#datefin').val().trim();
	var typeRecherche = $( "input[name*='filter']:checked" ).val();

	if (typeRecherche === "code") {
		idcoll = $('#idcoll').val().trim();
		equipe = $('#equipelabo').val().trim();;
	} else {
		idcoll = "";
		equipe = "";
		idshal = getIdsHal();
	}

	var soulign;
	if ($('#soulignauteur').is(':checked')) soulign=1; else soulign=0;
	// vidange des zones avant nouvel export : todo
	
	var datedeb;
	var laDate=new Date();
	iAnneeEnCours = laDate.getFullYear(); 
	iAnneeMax = laDate.getFullYear() + 1;
	
	if (dated === "") datedeb = 1900; else datedeb = parseInt(dated,10);
	
	var datefin;
	if (datef === "") datefin = iAnneeEnCours; else datefin = parseInt(datef,10);
	
	if (datedeb < 1900 ||  datedeb > iAnneeMax) {
		setError("Veuillez SVP saisir une date début valide.");
		return ;
	}
	if (datefin < 1900 ||  datefin > iAnneeMax) {
		setError("Veuillez SVP saisir une date de fin valide.");
		return ;
	}
	
	if (datedeb > datefin) {
		setError("La date de début doit être inférieure ou égale à la date de fin.");
		return ;
	}

	var words = idcoll.split(' ');
	if (words.length>1) {
		setError("Code de collection invalide : Veuillez saisir une chaine de caratères continue SVP.");
		return ;
	}

	if (typeRecherche === "aurehal" && idshal === "") {
		setError("Veuillez renseigner au moins un identifiant auréHAL.");
		return ;
	}


	$("#chargement").show(); // affichage de chargement en cours
	$("#ResultsHolder").slideUp();
	
	$.ajax({
		url: '/exporthal',
		type: 'POST',
		data: 'recherche=' + typeRecherche + '&idshal=' + idshal + '&idcoll=' + idcoll + '&dateDeb=' + datedeb + '&dateFin=' + datefin + '&equipelabo='+equipe + '&soulignauteur=0',
		dataType: 'JSON',
		success: function(data, status, xhr){
			$("#chargement").hide();
			if (data[2]!="Erreur") {
				$('#libelleBouton').show();
				$('#resultats').html(data[1]);            // #resultats : zone d'affichage des stats en partie droite
				$('#vhalbouton').attr("href", data[0]);   // #details : zone pour l'affichage des liens : API, lancer recherche sur Hal etc..
				$('#dlbouton').attr("href", data[2]);     // bouton de téléchargement du fichier rtf
				$('.nbreResultats').html(data[3]);
				$('#CodeCreaStats').val(data[4]);
				$('#dlbouton').show();

				CanvasJS.addCultureInfo("fr",
					{
						decimalSeparator: " ",
						digitGroupSeparator: ",",
						days: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
						months: ["Janvier", "Fevrier", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Decembre"],
						shortMonths: ["Janv.", "Fevr.", "Mars", "Avr.", "Mai", "Juin", "Juill.", "Aout", "Sept.", "Oct.", "Nov.", "Dec."],
					}
				);

				CanvasJS.addColorSet("SaphirColors", ["#ed145b","#e6155e","#df1660","#d71863","#d01966","#c91a69","#c21b6b","#ba1d6e","#b31e71","#ac1f73","#a52076","#9d2279","#96237b","#8f247e","#882581","#802783","#792886","#722989","#6b2a8b","#632c8e","#5c2d91"]);

				var chart = new CanvasJS.Chart("StatsChartContainer1", {
					backgroundColor: "transparent",
					culture: "fr",
					colorSet: "SaphirColors",
					exportEnabled: false,
					animationEnabled: true,
					legend:{
						fontFamily: "Source Sans Pro Light",
						fontColor: "#858795",
						fontSize: 12,
					},
					toolTip: {
						shared: true,
						borderColor: "#FFFFFF",
						fontColor: "#727482",
						fontSize: 13,
						fontFamily: "Source Sans Pro Light",
						cornerRadius: "4",
					},
					data: [{
						type: "pie",
						// innerRadius: 70,
						showInLegend: true,
						toolTipContent: "<b>{name}</b>: {y} (#percent%)",
						dataPoints: 
							eval(jQuery("#CodeCreaStats").val())
					}]
				});
				$("#ResultsHolder").slideDown();
				chart.render();

			} else {
				setError(data[1]);
			}
			//console.log(data);
			
		},
		error: function(jqXhr, textStatus, errorMessage){
			$("#chargement").hide();
			if (errorMessage=="Internal Server Error")
				setError("La recherche a échouée, veuillez contacter <a href=\"mailto:support-technique@hceres.fr\">support-technique@hceres.fr</a> si le problème persiste.");
			else
				setError(errorMessage);
			$('#resultats').show();
			console.log(textStatus);
		}
	});
}

let captcha;

$(document).ready(function(){
	const halCaptchaActif = document.getElementById('halCaptchaActif').dataset.myVariable;
	if(halCaptchaActif != 'none') {
		captcha = $('#botdetect-captcha').captcha({
			captchaEndpoint: 'captcha/simple-captcha-endpoint'
		});
	}
	
	M.updateTextFields();
	// au chargement de la page on cache le bouton de résultat
	$("#bouton_resultat").hide();
	$('#dlbouton').hide();
	$('#libelleBouton').hide();
	
	// on appuie sur la touche Entrée
	$('body').keypress(function(e){
		if( e.which == 13 ){
			if(halCaptchaActif != 'none') {
				var userEnteredCaptchaCode = captcha.getUserEnteredCaptchaCode();
				var captchaId = captcha.getCaptchaId();
				var postData = {
					code: userEnteredCaptchaCode,
					id: captchaId
				};
				var form = event;
				$.ajax({
					method: 'POST',
					url: 'captcha/validation',
					dataType: 'json',
					async: false,
					contentType: 'application/json; charset=utf-8',
					data: JSON.stringify(postData),
					success: function (response) {
						if (response == false) {
							// La validation a échoué. Le captcha est rechargé
							captcha.reloadImage();
							setError("Le CAPTCHA n'a pas été saisi correctement, refaites une tentative.");
							// form.preventDefault();
						} else {
							lance_recherche();
						}
					}
				});
			}else{
				lance_recherche();
			}
		}
	});
	
	
	// Lorsqu'on valide le formulaire
	$("#submitbouton").off("click").on("click", function(){
		if(halCaptchaActif != 'none') {
			var userEnteredCaptchaCode = captcha.getUserEnteredCaptchaCode();
			var captchaId = captcha.getCaptchaId();
			var postData = {
				code: userEnteredCaptchaCode,
				id: captchaId
			};
			var form = event;
			$.ajax({
				method: 'POST',
				url: 'captcha/validation',
				dataType: 'json',
				async: false,
				contentType: 'application/json; charset=utf-8',
				data: JSON.stringify(postData),
				success: function (response) {
					if (response == false) {
						// La validation a échoué. Le captcha est rechargé
						captcha.reloadImage();
						setError("Le CAPTCHA n'a pas été saisi correctement, refaites une tentative.");
						// form.preventDefault();
					} else {
						lance_recherche();
					}
				}
			});
		}else{
			lance_recherche();
		}
	}) 
		
	
	// clic sur le bouton ANNULER
	$("#cancelbouton").click(function(){
		$('#idcoll').val('');
		$('#datedeb').val('');
		$('#datefin').val('');
		$('#resultats').html(' ');
		$('#details').html(' ');
		$('#error-section').hide();
		$('#ids-hal').html(' ');
		$("#equipelabo").val('');
		$("#soulignauteur").prop("checked", false);
		$("#ResultsHolder").slideUp();
	})
	
	$("#chargement").hide();

	EnteteAttachHandler();

	$('.sidenav').sidenav();
	$('.parallax').parallax();
	$('.modal').modal();
	$('.tooltipped').tooltip();

	//activation / désactivation selon le boutton radio
	$( "input[name*='filter']" ).change(function(){
		var filter = $(this).val();
		if (filter === 'code') {
			$( "#aurehal" ).attr('disabled', 'disabled')
			$( "#equipelabo" ).removeAttr('disabled')
			$( "#idcoll" ).removeAttr('disabled')
		} else {
			$( "#aurehal" ).removeAttr('disabled')
			$( "#equipelabo" ).attr('disabled', true)
			$( "#idcoll" ).attr('disabled', true)
		}
	});

	//on clique sur ce bouton radio pour lancer l'événement change.
	$("#filter-code").click()

	//Click que le bouton ajouter id HAL pour créer les tags (chip)
	$( "#add-id-hal" ).click(function(){
		var content = $("#aurehal").val().trim();
		if (content !== "") {
			$("#ids-hal").append("<div class='chip' hal-value='"+content+"'>"+content+"<i class=\"close material-icons\">close</i></div>");
			$("#aurehal").val("");
		}
	});

	//Dans la zone de texte des ids, click sur le bouton "ajouter" lorsque que l'utilisateur appuie sur la touche entrer
	$('#aurehal').on('keypress',function(e) {
		if(e.which === 13) {
			$( "#add-id-hal" ).trigger('click');
			return false;
		}
	});

	$('#error-section').hide();





});
