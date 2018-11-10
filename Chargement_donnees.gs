/**
 * Met à jour le tableau "Strong Event" de la fiche en cours(dossier positionné le plus haut dans le "range actif") 
 * avec les données passées sous forme de fichier JSON.
 * @param {Object} fiche la spreadsheet à mettre à jour
 * @param {String} donnees les données à insérer
 */

function set_strong_event_fiche_current_range(donnees){
  try{
    var fiche=get_spreadsheet_from_current_range();
    fiche=fiche[0];
   set_strong_event_fiche(fiche,donnees);
   return "{\"message\":\"La fiche a été mise à jour\"}"
  }
  catch(e){
    var message="Aucun dossier existant choisi";
    gestion_erreur(e, message);
    return "{\"message\":\"Le tableau Strong Event n'a pas été mis à jour\"}"
  }
}
/**
 * Met à jour le tableau "Strong Event" de la fiche passée en paramètre avec les données passées
 * sous forme de fichier JSON.
 * @param {Object} fiche la spreadsheet à mettre à jour
 * @param {String} donnees les données à insérer
 */

function set_strong_event_fiche(fiche,donnees){
  try{
      donnees=JSON.parse(donnees);
    if(donnees['Date']==""){
      donnees['Date']=Utilities.formatDate(new Date(),"GMT+1", "yyyy-MM-dd' 'HH:mm:ss");
    }
    else if(donnees['Date'].match(/(J\+)([0-9]{1,2})/gi)!=null){
       donnees['Date']=calcul_date(parseInt(donnees['Date'].match(/([0-9]{1,2})/)));
    }
    else if(donnees['Date']="-"){
      donnees['Date']="";
    }
    else{
      throw "Problème avec le format de date";
    }
    var strong_event=fiche.getSheetByName("fiche").getRange(6,5,16,8);
    var strong_event_values=strong_event.getValues();
    var i=0;
  
    while((strong_event_values[i][0]!=donnees['Evenement'])&&(i<strong_event_values.length)){
     i++;
    }
    if((i==strong_event_values.length-1)&&(strong_event[i]!=donnees['Evenement'])){
     throw new Error('L\'événement n\'existe pas');
     }
     if(strong_event_values[i][1]==""){
      strong_event_values[i][1]=donnees['Date'];
      strong_event_values[i][2]=donnees['Type'];
      strong_event.setValues(strong_event_values);
     }
  }
  catch(e){
    var message="Aucun dossier existant choisi";
    gestion_erreur(e, message);
    return "{\"message\":\"Le tableau Strong Event n'a pas été mis à jour\"}"
  }

}

/**
 * Fonction permettant de récupérer les informations de l'onglet fiche.
 * Les 2 1ères colonnes (A et B) comportent des données "simples" donc il est possible de rajouter de nouvelles informations et qui 
 * seront prises en compte automatiquement par le script.
 * Les colonnes C  à H comportent des infos mixtes (des données simples + des tableaux double entrée)=> ne pas ajouter de données à ces colonnes
 * Tableau "Strong Event"=> possible de rajouter 10 lignes de données(hors libellés)
 * Tableaux "Vol initial" et "Vol de réacheminement" peuvent contenir jusque 4 vols chacun.
 *
 * @param {String} id l'id GAS de la fiche Spreadsheet
 * @return {String} La liste des passagers avec leurs infos en chaîne de caractères JSON.
 */

function get_fiche_dossier(id){
  try{
    //récupère l'ensemble des informations de l'onglet passager de la fiche + le nombre de passagers sur l'onglet fiche
    var fiche_dossier=SpreadsheetApp.openById(id).getSheetByName("fiche");

    var infos=fiche_dossier.getRange(1,1,fiche_dossier.getLastRow(),10).getValues(); 
    //variable permettant de constituer une chaîne de type JSON
    var string="{";
  
    //récupère les infos des 2 premières colonnes (colonnes de A à B)
    for(var i=0; i<infos.length;i++){
      for(var j=0; j<2;j=j+2){
        if(infos[i][j]!=""){
          string=string+"\""+infos[i][j]+"\":\""+format_for_json_string(infos[i][j+1])+"\",";
        }
      }
    } 
  
    //récupère les infos des colonnes de C à D, jusque la ligne 17
    for(var i=0; i<17;i++){
      for(var j=2; j<4;j=j+2){
        if(infos[i][j]!=""){
          string=string+"\""+infos[i][j]+"\":\""+format_for_json_string(infos[i][j+1])+"\","
        }
      }
    }
    
    //récupère les infos des colonnes de E à F, jusque la ligne 5
    for(var i=0; i<5;i++){
      for(var j=4; j<7;j=j+2){
        if(infos[i][j]!=""){
          string=string+"\""+infos[i][j]+"\":\""+format_for_json_string(infos[i][j+1])+"\","
        }
      }
    }
  
    //récupère les infos "Vol initial"
    string=string+"\"Vol initial\":[";
    if(infos[19][2]!=""){
      for(var i=19;i<23;i++){
        if(infos[i][2]!=""){
          string=string+"{";
          for(var j=2; j<8;j++){
            string=string+"\""+infos[18][j]+"\":\""+format_for_json_string(infos[i][j])+"\",";
          }
          string = string.substring(0,string.length-1);
          string += "},";
        }
      }
      string = string.substring(0,string.length-1);
    }
    string=string+"],";
  
    //récupère les infos "Vol de réacheminement"
    string=string+"\"Vol de réacheminement\":[";
    if(infos[24][2]!=""){
      for(var i=24;i<infos.length;i++){
        if(infos[i][2]!=""){
          string=string+"{";
          for(var j=2; j<8;j++){
            string=string+"\""+infos[18][j]+"\":\""+format_for_json_string(infos[i][j])+"\",";
          }
          string = string.substring(0,string.length-1);
          string += "},";
        }
      }
      string = string.substring(0,string.length-1);
    }
    string=string+"],";

    //récupère le tableau Strong event :  3 colonnes max et 10 lignes max de données (donc hors
    //libellés de colonnes et de lignes)
    string=string+"\""+infos[5][4]+"\":{";
    for(var i=6;i<16; i++){
      if(infos[i][4]!=""){
        string=string+"\""+infos[i][4]+"\":{";
        for(var j=5;j<8;j++){
          if(infos[5][j]!=""){
            string=string+"\""+infos[5][j]+"\":\""+format_for_json_string(infos[i][j])+"\",";
          }
        }
        string = string.substring(0,string.length-1);
        string += "},";
      }
    }
  

  string = string.substring(0,string.length-1);
  string += "}}";
  
  return string;
}
 catch(e){
      var message="Aucun dossier existant choisi";
      gestion_erreur(e, message);
    }
}


/**
 * Fonction permettant de mettre à jour les colonnes A et B de l'onglet "fiche".
 * Nécesssite une modification de la fonction get_infos_fiche pour récupérer les champs vides
 * 
 * @param {String} id l'id de la fiche
 * @param {String} donnees les données à mettre à jour
 * @return {String} message
 */
  
/*function set_infos_generale_fiche(id, donnees_entree){
  var fiche = SpreadsheetApp.openById(id);
  var donnees_entree = JSON.parse(donnees_entree);
  var infos_fiche = JSON.parse(get_passagers_dossier(id));
 
    for(var property in donnees_entree){
      infos_fiche[property]=donnees_entree[property];
    }

    
    var tableau_final=[];
    for(var i=0;i<Object.keys(donnees).length;i++){
      tableau_final[i][0]=
      for (var j=0;j<2;j++){
        if((donnees_maj[10*i+j]==undefined)||(donnees_maj[10*i+j].match(/champ_vide/)=="champ_vide")){
          tableau_final[i][j]="";
        }
        else{
          tableau_final[i][j]=donnees_maj[10*i+j];
        }
      }
    }
  fiche.getSheetByName("passagers").getRange(2,2,nb_passagers*6,10).setValues(tableau_final);

}*/

function set_ref_recla_current_range(num_reservation){
  try{
    var fiche=get_spreadsheet_from_current_range();
    fiche=fiche[0];
    if(fiche==""){throw{"message":"Dossier introuvable"};}
    set_ref_reclamation(num_reservation,fiche.getId());
  }
  catch(e){
      var message="Aucun dossier existant choisi";
      gestion_erreur(e, message);
  }
}
function set_ref_reclamation(num_reservation,id){
  var fiche = SpreadsheetApp.openById(id);
  fiche=fiche.getSheetByName("Fiche");
  fiche.getRange(15,2).setValue(num_reservation);
}

function set_circonstance_accident_current_range(circonstances_accicent){
  try{
    var fiche=get_spreadsheet_from_current_range();
    fiche=fiche[0];
    if(fiche==""){throw{"message":"Dossier introuvable"};}
    set_circonstance_accident(circonstances_accicent,fiche.getId());
  }
  catch(e){
      var message="Aucun dossier existant choisi";
      gestion_erreur(e, message);
  }
}
function set_circonstance_accident(circonstances_accicent,id){
  var fiche = SpreadsheetApp.openById(id);
  fiche=fiche.getSheetByName("Fiche");
  fiche.getRange(21,2).setValue(circonstances_accicent);
}


function get_passagers_dossier(id){
  try{
    var fiche = SpreadsheetApp.openById(id);
    //récupère l'ensemble des informations de l'onglet passager de la fiche + le nombre de passagers sur l'onglet fiche
    var passagers=fiche.getSheetByName("passagers");
    var liste_passagers=passagers.getRange(2,1,passagers.getLastRow()-1,11).getValues();
    var nb_passagers=fiche.getSheetByName("fiche").getRange(25,2).getValue();
 
    //variable permettant de constituer une chaîne de type JSON => basé sur 6 lignes par passager.
    var string="{";

    //variable permettant de générer un nom différent pour chaque champ vide
    //champs vides nécessaires pour la fonction set_infos_passagers
    var n=0
  
    for(var i=0;i<6*nb_passagers;i=i+6){
      string=string+"\""+liste_passagers[i][0]+"\":{";
      for(var j=i;j<i+6;j++){
        for(var k=1;k<liste_passagers[j].length;k=k+2){
         if(liste_passagers[j][k]!=""){
            string=string+"\""+liste_passagers[j][k]+"\":\""+liste_passagers[j][k+1]+"\",";
         }
         else{
          string=string+"\""+"champ_vide"+n+"\":\"\",";
          n++;
         }
        }
      }
      string = string.substring(0,string.length-1);
      string += "},";
  }
    string = string.substring(0,string.length-1);
    string=string+"}";
    return string;
}
   catch(e){
      var message="Impossible de récupérer les données de la fiche" + id;
      gestion_erreur(e, message);
    }
}

/**
 * Fonction permettant de mettre à jour les infos passager de l'onglet "passagers".
 * La fonction est basée sur 6 lignes par passagers dans l'onglet "passagers".
 * Le nombre de passagers est récupéré sur l'onglet fiche.
 *
 * @param {String} id l'id de la fiche
 * @param {String} donnees les données à mettre à jour
 * @return {String} message
 */
  
 function set_infos_passagers(id,donnees_entree){
  var fiche = SpreadsheetApp.openById(id);
  var donnees_entree = JSON.parse(donnees_entree);
  var passagers = JSON.parse(get_passagers_dossier(id));

  for(var passager in donnees_entree){ 
    for(var property in donnees_entree[passager]){
      passagers[passager][property]=donnees_entree[passager][property];
    }
  }

  var nb_passagers=Object.keys(passagers);
  nb_passagers=nb_passagers.length;
  var donnees_maj=[];
  var n=0;
  for(passager in passagers){
    for(property in passagers[passager]){
      donnees_maj[n]=property;
      donnees_maj[n+1]=passagers[passager][property];
      n=n+2;
    }
  }

  var tableau_final=[];
    for(var i=0;i<nb_passagers*6;i++){
      tableau_final[i]=[];
      for (var j=0;j<10;j++){
        if((donnees_maj[10*i+j]==undefined)||(donnees_maj[10*i+j].match(/champ_vide/)=="champ_vide")){
          tableau_final[i][j]="";
        }
        else{
          tableau_final[i][j]=donnees_maj[10*i+j];
        }
      }
    }
  fiche.getSheetByName("passagers").getRange(2,2,nb_passagers*6,10).setValues(tableau_final);
 }

/**
 * Fonction permettant de récupérer les infos d'un tableau sous forme de chaîne de caractères JSON
 *
 * @param {String} id : l'id du fichier
 * @param {String} name : le nom de l'onglet
 * @return {String} Les infos de la fiche en chaîne de caractères JSON.
 */

function get_infos_tableau(id, name){
  try{
    //ouvre la fiche et l'onglet et récupère l'ensemble des infos
    var spreadsheet = SpreadsheetApp.openById(id);
    var sheet = spreadsheet.getSheetByName(name);
    var liste = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn()).getValues();
  
    var string="{\"données\": [";
  
    //transformation en objet json
    for(var i=1; i<liste.length;i++){
      string=string+"{";
      for(var j=0; j<liste[i].length;j++){
        string=string+"\""+liste[0][j].trim()+"\":\""+liste[i][j]+"\",";
      }
      string=string.substring(0,string.length-1);
      string=string+"},";
    }
  
    string=string.substring(0,string.length-1);
    string=string+"]}";
    return string;
  } 
   catch(e){
      var message="Impossible de récupérer les informations demandées";
      gestion_erreur(e, message);
      return "{\"message\":\""+message+"\"}"
    }
}

/**
 * Renvoie les infos compagnie du dossier sous forme de chaîne de caractères JSON (dossier positionné
 * le plus haut dans le "range actif")
 * @return {String} Les infos de la fiche en chaîne de caractères JSON.
 */

function get_infos_compagnie_from_current_range(){
  try{
    var ligne_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dossiers').getActiveRange().getRow();
    var compagnie_fiche= SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dossiers').getRange(ligne_spreadsheet,14).getValue();
    var compagnie;
    //récupère les infos de la compagnie
    var infos_compagnies_ss = data_airlines_airport.getSheetByName('airlines');
    var infos_compagnies = infos_compagnies_ss.getRange(1,1,infos_compagnies_ss.getLastRow(),infos_compagnies_ss.getLastColumn()).getValues();

    var i = 0;
    while((infos_compagnies[i][3]!=compagnie_fiche)&(i<infos_compagnies.length)){
      i++;
    }
    var string="{";
      for(var j=0; j<infos_compagnies[i].length;j++){
        string=string+"\""+infos_compagnies[0][j]+"\":\""+infos_compagnies[i][j]+"\",";
      }
      string=string.substring(0,string.length-1);
      string=string+"}";
      var aremplacer ="\n";
      var re = new RegExp(aremplacer,'g');
      string=string.replace(re,"<br />");
    return string;
  }
  catch(e){
    var message="Impossible de récupérer les informations demandées";
    gestion_erreur(e, message);
    return "{\"message\":\""+message+"\"}"
  }
}