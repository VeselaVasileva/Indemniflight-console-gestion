function archiver_dossiers_current_range(arguments){
    try{
        var fiches=get_spreadsheet_from_current_range();
        
        arguments=JSON.parse(arguments);
        var motif=arguments['motif'];
        var commentaire=arguments['commentaire'];
        var type_email="{\"name\":\""+arguments['name']+"\",\"modifiable\":"+arguments['modifiable']+"}"

        for(var i=0;i<fiches.length;i++){
            if((fiches[i]==undefined)||(fiches[i]=="")||(fiches[i]==[])){throw {"message":"Dossier introuvable"};}
            archiver_dossier(fiches[i],motif,commentaire);
                //Envoi l'email au client si la case est cochée
                if(arguments['envoi_email_client']==true){
                    generer_emails_gas_gestion_doc(type_email);
                }
        }
        return "{\"message\":\"Dossier archivé avec succès\"}"
    }

    catch(e){
        var message="";
        gestion_erreur(e, message);
        return "{\"message\":\"Une erreur s'est produite lors de l'archivage du dossier : "+e.message+"\"}";
      }

}



function archiver_dossier(fiche,motif,commentaire){ 
    var id=fiche.getId();
    //Met à jour le tracer
    var tracer="{\"type\":\"tracer\",\"id\":\""+id+"\",\"copie\":0, \"contenu\":[{\"Type event\":\"Interne\",\"Evenement\":\"Archive dossier\",\"Date\":\"\",\"Prochain évènement\":\"-\",\"Date prévue\":\"-\",\"Statut dossier\":\"Dossier archivé : "+motif+"\",\"Document relatif\":\"\",\"Canal\":\"\",\"Commentaire\":\""+commentaire+"\"}]}";
    set_tracer(tracer);

    //Met à jour le tracer avocat si ce dernier existe
    if(fiche.getSheetByName("tracer_avocat")!=undefined){
        tracer="{\"type\":\"tracer_avocat\",\"id\":\""+id+"\",\"copie\":0, \"contenu\":[{\"Evenement\":\"Archive dossier\",\"Date\":\"\",\"Prochain évènement\":\"-\",\"Date prévue\":\"-\",\"Statut dossier\":\"Dossier archivé"+motif+"\",\"Commentaire\":\"\"}]}";
        set_tracer(tracer);
    }
 

    //Supprime le dossier des dossiers en cours
    var dossier=get_dossier_from_name(fiche.getName());
    bdd_dossiers_archives.addFolder(dossier);
    bdd_dossiers_en_cours.removeFolder(dossier);

    //vérifie si le dossier a d'autres dossiers parents
    var dossiers_parents=dossier.getParents();
    if(dossiers_parents.hasNext()){
        var dossiers_avocats=get_infos_tableau(data_avocats,"liste_avocats");
        dossiers_avocats=JSON.parse(dossiers_avocats);

        //si le dossier a autres dossiers parents et ils correspondent à un id dossier avocat,
        //le dossier est déplacé dans les archives de cet avocat
        while(dossiers_parents.hasNext()){
            var id_dossier_parent=dossiers_parents.next().getId();
            for(var i=0;i<dossiers_avocats['données'].length;i++){
                if(dossiers_avocats['données'][i]['Dossier en cours']==id_dossier_parent){
                    (DriveApp.getFolderById(dossiers_avocats['données'][i]['Dossier en cours'])).removeFolder(dossier);
                    (DriveApp.getFolderById(dossiers_avocats['données'][i]['Dossier Archivé'])).addFolder(dossier)
                }
            }
        }
    }

    //Met à jour la fiche
    var note_archive = fiche.getSheetByName("fiche").getRange(5,4);
    note_archive.setValue("Archivé");
    note_archive.setBackground("#ffa500");
    note_archive.setFontSize(16);
    note_archive.setFontColor("#ffffff");

    //Met à jour le tableau récap des dossiers en cours
    var liste_dossiers_existants_values=liste_dossiers_existant.getRange(1,1,liste_dossiers_existant.getLastRow(),liste_dossiers_existant.getLastColumn()).getValues();
    var i=0;
    while((liste_dossiers_existants_values[i][2]!=id)&&(i<liste_dossiers_existants_values.length)){
        i++;
    }
    if((i==liste_dossiers_existants_values.length)&&(liste_dossiers_existants_values[i][2]!=id)){
        throw "{\"message\":\"dossier introuvable\"}";
    }

    liste_dossiers_existant.getRange(i+1,4).setValue("true");

    return "{\"message\":\"Dossier archivé avec succès\"}"
}


function desarchiver_dossiers_current_range(arguments){
    try{
        var fiches=get_spreadsheet_from_current_range();
        
        arguments=JSON.parse(arguments);
        var motif=arguments['motif'];
        var commentaire=arguments['commentaire'];

        for(var i=0;i<fiches.length;i++){
            if((fiches[i]==undefined)||(fiches[i]=="")||(fiches[i]==[])){throw {"message":"Dossier introuvable"};}
            desarchiver_dossier(fiches[i],motif,commentaire);
        }
        return "{\"message\":\"Dossier désarchivé avec succès\"}"
    }

    catch(e){
        var message="";
        gestion_erreur(e, message);
        return "{\"message\":\"Erreur : "+e.message+"\"}";
      }

}



function desarchiver_dossier(fiche,motif,commentaire){  
    var id=fiche.getId();
    //Met à jour le tracer
    var num_row=fiche.getSheetByName("tracer").getLastRow()-1;
    var tracer="{\"type\":\"tracer\",\"id\":\""+id+"\",\"copie\":0, \"contenu\":[{\"Type event\":\"Interne\",\"Evenement\":\"Dossier désarchivé : "+motif+"\",\"Date\":\"\",\"Prochain évènement\":\"-\",\"Date prévue\":\"-\",\"Statut dossier\":\"-\",\"Document relatif\":\"\",\"Canal\":\"\",\"Commentaire\":\""+commentaire+"\"},"+tracer_get_row(id,"tracer",num_row)+"]}";
    set_tracer(tracer);

    if(fiche.getSheetByName("tracer_avocat")!=undefined){
        //Met à jour le tracer avocat
        num_row=fiche.getSheetByName("tracer_avocat").getLastRow()-1;
        tracer="{\"type\":\"tracer_avocat\",\"id\":\""+id+"\",\"copie\":0, \"contenu\":[{\"Evenement\":\"Dossier désarchivé\",\"Date\":\"\",\"Prochain évènement\":\"-\",\"Date prévue\":\"-\",\"Statut dossier\":\"-\",\"Commentaire\":\"\"},"+tracer_get_row(id,"tracer_avocat",num_row)+"]}";
        set_tracer(tracer);
    }
    
    //Met à jour la fiche
    var note_archive = fiche.getSheetByName("fiche").getRange(5,4);
    note_archive.setValue("");
    note_archive.setBackground("#ffffff");
    note_archive.setFontSize(10);
    note_archive.setFontColor("#000000");


    //Supprime le dossier des dossiers archivés
    var dossier=get_dossier_from_name(fiche.getName());
    bdd_dossiers_archives.removeFolder(dossier);

    //vérifie si le dossier a d'autres dossiers parents
    var dossiers_parents=dossier.getParents();
    if(dossiers_parents.hasNext()){
        var dossiers_avocats=get_infos_tableau(data_avocats,"liste_avocats");
        dossiers_avocats=JSON.parse(dossiers_avocats);

        //si le dossier a autres dossiers parents et ils correspondent à un id dossier avocat,
        //le dossier est déplacé dans les archives de cet avocat
        while(dossiers_parents.hasNext()){
            var id_dossier_parent=dossiers_parents.next().getId();
            for(var i=0;i<dossiers_avocats['données'].length;i++){
                if(dossiers_avocats['données'][i]['Dossier Archivé']==id_dossier_parent){
                    (DriveApp.getFolderById(dossiers_avocats['données'][i]['Dossier Archivé'])).removeFolder(dossier);
                    (DriveApp.getFolderById(dossiers_avocats['données'][i]['Dossier en cours'])).addFolder(dossier)
                }
            }
        }
    }

    bdd_dossiers_en_cours.addFolder(dossier);

    //Met à jour le tableau récap des dossiers en cours
    var liste_dossiers_existants_values=liste_dossiers_existant.getRange(1,1,liste_dossiers_existant.getLastRow(),liste_dossiers_existant.getLastColumn()).getValues();
    var i=0;
    while((liste_dossiers_existants_values[i][2]!=id)&&(i<liste_dossiers_existants_values.length)){
        i++;
    }
    if((i==liste_dossiers_existants_values.length)&&(liste_dossiers_existants_values[i][2]!=id)){
        throw "{\"message\":\"dossier introuvable\"}";
    }

    liste_dossiers_existant.getRange(i+1,4).setValue(false);
}