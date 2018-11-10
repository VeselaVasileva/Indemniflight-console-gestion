function generer_facture_current_range(){
  try{
    var fiche = get_spreadsheet_from_current_range();
    fiche=fiche[0];
    generer_facture(fiche);
    return "{\"message\":\"Facture générée avec succès\"}"
  }
  catch(e){
    var message="Vérifiez si le courrier de réclamation existe";
    gestion_erreur(e, message);
    return "{\"message\":\"Problème avec la génération de la facture\"}"
  }
}


function generer_facture(fiche){
  var dossier_parent=DriveApp.getFolderById("1hcmgm0eRyTnsyzjXqVqDFARRbaVFiOew");
  var annee=new Date();
  annee=annee.getFullYear();
  var donnees = get_fiche_dossier(fiche.getId());

  //vérifier si le dossier Factures de l'année en cours existe et le crée si besoin
  if(dossier_parent.getFoldersByName("Factures_"+annee).hasNext()){
    var dossier=dossier_parent.getFoldersByName("Factures_"+annee).next();
  }
  else{
    var dossier=dossier_parent.createFolder("Factures_"+annee);
  }
  
    //Création de la facture
    var facture=DriveApp.getFileById(template_facture).makeCopy("Facture",dossier);
    var document=DocumentApp.openById(facture.getId());
    var body=document.getBody();
    //var corps=DocumentApp.openById(facture.getId()).getText();
    body = publipostage_body(body, donnees);
    var compte_cantonnement_factures=compte_cantonnement.getSheetByName("Facture");
    var num_facture=compte_cantonnement_factures.getRange(compte_cantonnement_factures.getLastRow(),1).getValue()+1;

    donnees=JSON.parse(donnees);
    
    if((donnees['Commission']<0)||(donnees['Commission']>0.30)){throw {"message":"Problème avec la comission"};}
    if(donnees['Commission'] == "" || donnees['Commission'] == 0){donnees['Commission'] = 0.25; }
  
    var TTC=donnees['Commission']*donnees['Indemnisation totale'];
    var HT=Math.ceil(TTC / 1.2 * 100) / 100;
    body=body.replaceText("#montant_ttc#", TTC);

    //var montant_ht=new RegExp('#montant_ht#','g');
    body=body.replaceText("#montant_ht#", HT);
    body=body.replaceText("#tva#", Math.ceil((TTC - Math.ceil(TTC / 1.2 * 100) / 100) * 100) / 100);
    body=body.replaceText("#num_fact#",num_facture)
    //body=body.setText(corps);
   document.saveAndClose();

    document.setName(annee+"_"+donnees['Numéro du dossier']+"_"+num_facture);
    var newFile = dossier.createFile(document.getAs('application/pdf'));
    facture.setTrashed(true);
    var date=Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd' 'HH:mm:ss");
    var ligne =[num_facture,date,donnees['Numéro du dossier'],TTC["montant_fact_ht"],TTC["montant_fact_ttc"],"","","",""];
    var compte_cantonnement_values=compte_cantonnement_factures.getRange(1,1,compte_cantonnement_factures.getLastRow(), compte_cantonnement_factures.getLastColumn()).getValues();
    compte_cantonnement_values.push(ligne);
    compte_cantonnement_factures.getRange(1,1,compte_cantonnement_values.length,compte_cantonnement_factures.getLastColumn()).setValues(compte_cantonnement_values);
    return "{\"message\":\"La facture a été généré\"}"

}



/**
 * Récupère le texte de la lettre amiable pour le dossier en cours en fr ou en ou fr et en
 *
 * @param {String} langue la langue choisie dans la console.
 * @return {String} le texte de la lettre amiable.
 */  
function copier_courrier_amiable_current_range(langue){
  try{
    var fiche = get_spreadsheet_from_current_range();
    fiche=fiche[0];
    if((fiche=="")||(fiche==undefined)){throw{"message":"Dossier introuvable"};}
    return copier_courrier_amiable(fiche,langue);
  }
  catch(e){
    var message="Vérifiez si le courrier de réclamation existe";
    gestion_erreur(e, message);
    return "{\"message\":\"Erreur : "+e.message+"\"}"
  }
}

/**
 * Récupère le texte de la lettre amiable pour le dossier de la fiche passée en paramètre en fr ou en ou fr et en
 *
 * @param {String} langue la langue choisie (langue par défaut fr)
 * @return {String} le texte de la lettre amiable.
 */  

function copier_courrier_amiable(fiche,langue){

    if(get_dossier_from_name(fiche.getName()).getFoldersByName("Courrier_reclamation").hasNext()){
      var dossier=get_dossier_from_name(fiche.getName()).getFoldersByName("Courrier_reclamation").next()
    }
    else{
      throw {"message":"Le dossier Courrier de réclamation n'existe pas"};
    }

    if(dossier.getFilesByName("courrier_reclamation").hasNext()){
      var file=dossier.getFilesByName("courrier_reclamation").next();
    }
    else{
      throw {"message":"Le dossuer Courrier de réclamation est vide"};
    }
    
    var corps=DocumentApp.openById(file.getId()).getBody().getText();

    var Tablebody=corps.split("Executive Officer\nIndemniflight");
      switch(langue){
      case 'fr':
        var TexteLg=Tablebody[0]+'Executive Officer\nIndemniflight';
        break;
      case 'en':
        var TexteLg=Tablebody[1]+'Executive Officer\nIndemniflight';
        TexteLg=TexteLg.slice(4,TexteLg.length);
        break;
      case 'fr-en':
        var TexteLg=Tablebody[0]+'Executive Officer\nIndemniflight'+Tablebody[1]+'Executive Officer\nIndemniflight';
        break;
      default :
        var TexteLg=Tablebody[0]+'Executive Officer\nIndemniflight';
      }


    return TexteLg
}

/**
 * Génère la lettre amiable pour le dossier choisi dans la console en 3 formats : doc, html, pdf.
 *
 * @return {String} le message à afficher dans la console
 */  

function generer_lettre_amiable_current_range() {
  try{
    var dossier=get_spreadsheet_from_current_range();
    dossier=dossier[0];
    if(dossier==""){throw{"message":"Dossier introuvable"};}
    generer_lettre_amiable(dossier);
    return "{\"message\":\"Le courrier a été regénéré\"}"
  }
  catch(e){
    var message="";
    gestion_erreur(e, message);
    return "{\"message\":\"Erreur : "+e.message+"\"}";
  }
}

/**
 * Génère la lettre amiable pour la fiche passée en paramètre en 3 formats : doc, html, pdf.
 *
 * @return {String} le message à afficher dans la console
 */ 

function generer_lettre_amiable(fiche) {
    //Créer un dossier et récupérer son ID
    var sous_dossier = create_new_folder(fiche.getName(),"Courrier_reclamation");
    
    //Récupérer les données de la fiche
    var donnees = get_fiche_dossier(fiche.getId());
    var donnees_parse=JSON.parse(donnees);
    var anomalie=donnees_parse['Anomalie'];
    anomalie=anomalie.substring(anomalie.indexOf(" ")+1,anomalie.length);
    //Récupérer les données de l'onglet passagers et établir la liste des passagers
    var passagers = JSON.parse(get_passagers_dossier(fiche.getId()));
    var liste_passagers ="";
    for(var passager in passagers){ 
      liste_passagers=liste_passagers+format_name_passenger(passagers[passager]['Prénom'])+" "+format_name_passenger(passagers[passager]['Nom'])+", ";
    }
    liste_passagers=liste_passagers.substring(0,liste_passagers.length-2);
    
    //ajouter la liste des passagers aux arguments de la fonction publipostage
    donnees=donnees.substring(0,donnees.length-1)+", \"passagers\":\""+liste_passagers+"\"}";
    
    if(anomalie==""){
      throw {"message":"Pas d'anomalie signalée dans la fiche"}
    }
    //Met à jour le template html
    if(map_templates["courrier_reclamation_"+anomalie+"_html"]){
    var template_html=DriveApp.getFileById(map_templates["courrier_reclamation_"+anomalie+"_html"]).makeCopy("courrier_reclamation_html",sous_dossier);
    }else
    {
      throw {"message":"Pas de courrier de réclamation type pour cette anomalie"}
    }
    template_html=DocumentApp.openById(template_html.getId());
    var body_html=template_html.getBody();
    //var corps_html=body_html.getText();
    body_html=publipostage_body(body_html, donnees);
   // body_html=body_html.setText(corps_html);
    template_html.saveAndClose();
    
    
    //Met à jour le document et crée le pdf
    if(map_templates["courrier_reclamation_"+anomalie]){
      var template=DriveApp.getFileById(map_templates["courrier_reclamation_"+anomalie]).makeCopy("courrier_reclamation",sous_dossier);
    }
    else{ 
      throw {"message":"Pas de courrier de réclamation type pour cette anomalie"}
    }
    template=DocumentApp.openById(template.getId());
    var body=template.getBody();
    //var corps=body.getText();
    body = publipostage_body(body, donnees);
    //body=body.setText(corps);
    template.saveAndClose();
    var newFile = sous_dossier.createFile(template.getAs('application/pdf'));
}


/**
 * Génère les mandats pour tous les passagers du dossier choisi dans la console en 2 formats : doc et pdf.
 *
 * @return {String} le message à afficher dans la console
 */  
function generer_mandats_current_range(){
  try{
    var spreadsheet_dossier=get_spreadsheet_from_current_range();
    spreadsheet_dossier=spreadsheet_dossier[0];
    if(spreadsheet_dossier==""){throw{"message":"Dossier introuvable"};}
    //var ligne_spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dossiers').getActiveRange().getRow();
    //var spreadsheet_dossier = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dossiers').getRange(ligne_spreadsheet,1).getValue();
    //spreadsheet_dossier = get_spreadsheet_from_name(spreadsheet_dossier);
    generer_mandats(spreadsheet_dossier);
    return "{\"message\":\"Mandats regénérés\"}";
  }
  catch(e){
    var message="";
    gestion_erreur(e, message);
    return "{\"message\":\"Erreur : "+e.message+"\"}";
  }
}

/**
 * Génère les mandats pour tous les passagers du dossier passé en paramètre en 2 formats : doc et pdf.
 *
 * @param {Object} fiche la spreadsheet
 * @return {String} le message à afficher dans la console
 */ 

function generer_mandats(fiche){
    //Créer un dossier et récupérer son ID
    var sous_dossier = create_new_folder(fiche.getName(),"Mandats");
    
    //Récupérer les données de la fiche
    var donneesJSON = get_fiche_dossier(fiche.getId());
    var donnees = JSON.parse(donneesJSON);
    var langue = donnees['Langue'];
    
    //Récupérer les données de l'onglet passagers
    var passagers = JSON.parse(get_passagers_dossier(fiche.getId()));
    
    //Générer les mandats pour chaque passager
    for(var passager in passagers){       
      //mise à jour du mandat manuscrit
      var template_file=DriveApp.getFileById(map_templates["Mandat_template_"+langue]).makeCopy("Mandat_"+passagers[passager]['Prénom']+" "+passagers[passager]['Nom'],sous_dossier);
      var template=DocumentApp.openById(template_file.getId());
      var body=template.getBody();
      //var corps=body.getText();
      
      body=publipostage_body(body,JSON.stringify(passagers[passager]));
      body=publipostage_body(body,donneesJSON);

      template.saveAndClose();
      var newFile = sous_dossier.createFile(template.getAs('application/pdf'));
      template_file.setTrashed(true);  
  }
}