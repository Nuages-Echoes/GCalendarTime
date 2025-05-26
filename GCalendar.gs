function lireCalendrier() {
  // On récupère les calendriers de l'utilisateur actif 
  var calendriers = CalendarApp.getAllOwnedCalendars();
  var calnom;
  var calT; // calendrier de travail
  var evPeriode; // liste d'évènements sur la periode 
  let sheet = SpreadsheetApp.getActiveSheet();
  var tempsTravail;
  var tempsTravailNuit;
  var evNom;

  // lire les dates de début et de fin de période
  var dateD = sheet.getRange("A2").getValue();
  var dateF = sheet.getRange("B2").getValue();


  sheet.appendRow(["Nom du Salarié","Heures de jour","Heures de nuit","Heures totales","Temps de trajet"])
  for (let x = 0; x< calendriers.length; x ++){
    tempsTravail = 0;
    tempsTravailNuit = 0;
    tempsTrajet = 0;
    calnom = calendriers[x].getName();
    calT = calendriers[x];
    // heure de fin et de début du travail de nuit
    var nuitF = sheet.getRange("F3").getValue();
    var nuitD = sheet.getRange("E3").getValue();

    // récupérer les événements sur la période
    evPeriode = calT.getEvents(dateD,dateF);
    for (let y = 0; y< evPeriode.length; y ++){
      evNom = evPeriode[y].getTitle();
      debutT = evPeriode[y].getStartTime();
      finT = evPeriode[y].getEndTime();
      dureeEvent = calculerTemps(debutT,finT);
      tempsTravail= tempsTravail + dureeEvent;
      if((debutT.getHours() < nuitF) || (finT.getHours() >= nuitD) || (finT.getHours() == 0) || (debutT.getHours() >= nuitD) || ((finT.getHours() >= nuitD) && (finT.getMinutes != 0))){
        tempsTravailNuit= tempsTravailNuit + calculerTempsNuit_V2(debutT,finT,sheet);
      }
      if(evPeriode[y].getColor() == CalendarApp.EventColor.GRAY){
        tempsTrajet= tempsTrajet + dureeEvent;
      }

    }
    
    // affiche un ligne contenant le nom du calendrier, le temps travaille de jour, le temps travaille de nuit
    sheet.appendRow([calnom,tempsTravailleHM(tempsTravail-tempsTravailNuit),tempsTravailleHM(tempsTravailNuit),tempsTravailleHM(tempsTravail),tempsTravailleHM(tempsTrajet)])
  }

}

// cette fonction renvoie le temps en minutes entre deux objets Date au format
// 
function calculerTemps(dateD, dateF) {
  var heureD = dateD.getHours();
  var heureF = dateF.getHours();
  var minD = dateD.getMinutes();
  var minF = dateF.getMinutes();
  var nbHeures = 0;
  var nbMinutes = 0;
  var tempsEnMinutes = 0;

  if(dateD.getDay() == dateF.getDay()){
    if(minD <= minF){
      nbHeures= heureF-heureD;
      nbMinutes = minF - minD;
    } else { // si les minutes de début sont supérieures au temps en minute de fin
      nbHeures= heureF-heureD-1;
      nbMinutes = minF - minD + 60;
    }
  } else {
    nbHeures = 24-heureD+heureF;
    if(minD <= minF){
      nbMinutes = minF - minD;
    } else { // si les minutes de début sont supérieures au temps en minute de fin
      nbMinutes = minF - minD;
    }
  }

  tempsEnMinutes = nbMinutes + (nbHeures * 60);

  return tempsEnMinutes;
}

function calculerTempsNuit(dateD, dateF, sheet){
  var minutesDeNuit = 0
  var heureD = dateD.getHours();
  var heureF = dateF.getHours();
  var minD = dateD.getMinutes();
  var minF = dateF.getMinutes();
  var nbHeures = 0;
  var nbMinutes = 0;
  // heure de début travail de nuit
  var nuitD = sheet.getRange("E3").getValue();
  // heure de fin du travail de nuit
  var nuitF = sheet.getRange("F3").getValue();

  // si l'évènement termine après 21H00
  if((heureF >= nuitD) && (minF > 0)){
      // il faut comptabiliser les heures de nuit si une partie de l'évènement se passe pendant la nuit
    if(heureD < nuitD && heureD > nuitF) {heureD = nuitD};
    if(heureF > nuitF) {heureF = nuitF};

    if(dateD.getDay() == dateF.getDay()){
      if(minD <= minF){
        nbHeures= heureF-heureD;
        nbMinutes = minF - minD;
      } else { // si les minutes de début sont supérieures au temps en minute de fin
        nbHeures= heureF-heureD-1;
        nbMinutes = minF - minD + 60;
      }
    } else {
      nbHeures = 24-heureD+heureF;
      if(minD <= minF){
        nbMinutes = minF - minD;
      } else { // si les minutes de début sont supérieures au temps en minute de fin
        nbMinutes = minF - minD + 60;
      }
    }

  }
  
  if(heureF == nuitF) {nbMinutes = nbMinutes - minF}
  minutesDeNuit = nbMinutes + (nbHeures * 60);

  return minutesDeNuit;

}

function calculerTempsNuit_V2(dateD, dateF, sheet){
  var minutesDeNuit = 0
  var heureD = dateD.getHours();
  var heureF = dateF.getHours();
  var minD = dateD.getMinutes();
  var minF = dateF.getMinutes();
  var nbHeures = 0;
  var nbMinutes = 0;
  // heure de début travail de nuit
  var nuitD = sheet.getRange("E3").getValue();
  // heure de fin du travail de nuit
  var nuitF = sheet.getRange("F3").getValue();

  // si l'évènement commence avant 21H (mais après 6H) mais termine après 21H
  if((heureD < 21) && (heureD > 6)){
    heureD = 21;
    minD = 0;
  }

  // si l'évènement commence après minuit mais avant 6H
  if(heureF >= 6 && heureF < 21){
    heureF = 6;
    minF = 0;
  }

  if(dateD.getDay() == dateF.getDay()){
    if(minD <= minF){
      nbHeures= heureF-heureD;
      nbMinutes = minF - minD;
    } else { // si les minutes de début sont supérieures au temps en minute de fin
      nbHeures= heureF-heureD-1;
      nbMinutes = minF - minD + 60;
    }
  } else {
    nbHeures = 23-heureD+heureF;
    if(minD <= minF){
      nbMinutes = minF - minD;
    } else { // si les minutes de début sont supérieures au temps en minute de fin
      nbMinutes = minF - minD + 60;
    }
  }
  
  if(heureF == nuitF) {nbMinutes = nbMinutes - minF}
  minutesDeNuit = nbMinutes + (nbHeures * 60);

  return minutesDeNuit;

}

// renvoie un chaîne de caractères sous le format nb d'heures travaillés et minutes en centiemes
function tempsTravailleHM(tempsEnMinutes){
  var temps="";
  var heuresTravaillees=0;
  var minutesTravaillees=0;

  heuresTravaillees=Math.floor(tempsEnMinutes/60);
  minutesTravaillees=((tempsEnMinutes%60)*100)/60;
  if(minutesTravaillees > 10){
    temps=temps.concat(heuresTravaillees,",",Math.round(minutesTravaillees));
  } else {
    temps=temps.concat(heuresTravaillees,",0",Math.round(minutesTravaillees));
  }

  return temps;
}
