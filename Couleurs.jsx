#target illustrator

// Désactiver les alertes utilisateur pour éviter les interruptions
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

// Afficher un message de confirmation pour vérifier la taille du plan de travail
var userConfirmed = confirm("Avez-vous vérifié que la taille du plan de travail du document template correspond aux fichiers que vous allez traiter ? Cliquez sur Oui pour continuer ou Non pour arrêter.");

if (userConfirmed) {
    // Demander à l'utilisateur de sélectionner le dossier contenant les fichiers Illustrator à traiter
    var folder = Folder.selectDialog("Sélectionnez le dossier contenant les fichiers Illustrator");

    if (folder != null) {
        var files = folder.getFiles("*.ai");

        // Assurer qu'il y a au moins un fichier .ai dans le dossier
        if (files.length > 0) {
            // Récupérer le document template actuellement ouvert
            var templateDoc = app.activeDocument;

            // Désactiver à nouveau les alertes utilisateur pour éviter les interruptions
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

            for (var i = 0; i < files.length; i++) {
                // Ouvrir le fichier .ai
                var doc = app.open(files[i]);

                // Sélectionner tous les objets dans le fichier ouvert
                doc.selectObjectsOnActiveArtboard();
                app.copy();

                // Activer le document template
                app.activeDocument = templateDoc;

                // Coller les objets dans le document template
                app.executeMenuCommand('pasteInPlace');

                // Fusionner les couleurs globales
                var docSwatches = doc.swatches;
                var templateSwatches = templateDoc.swatches;

                for (var j = 0; j < docSwatches.length; j++) {
                    var swatch = docSwatches[j];

                    if (swatch.colorType == ColorModel.SPOT || swatch.colorType == ColorModel.PROCESS) {
                        try {
                            var templateSwatch = templateSwatches.getByName(swatch.name);
                            swatch.color = templateSwatch.color;
                        } catch (e) {
                            // Si la couleur n'existe pas dans le template, continuez
                        }
                    }
                }

                // Enregistrer le template avec le nom du fichier .ai
                var saveFile = new File(folder + '/' + files[i].name);
                templateDoc.saveAs(saveFile);

                // Réinitialiser le document template (supprimer les objets collés)
                templateDoc.activeLayer.pageItems.removeAll();

                // Fermer le fichier .ai sans sauvegarder
                doc.close(SaveOptions.DONOTSAVECHANGES);
            }

            // Réactiver les alertes utilisateur après traitement
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;

            // Afficher un message indiquant que tous les fichiers ont été traités
            alert("Tous les fichiers ont été traités avec succès.");

            // Fermer le dernier fichier traité (le templateDoc qui est maintenant enregistré)
            templateDoc.close(SaveOptions.DONOTSAVECHANGES);
        } else {
            alert("Aucun fichier .ai trouvé dans le dossier sélectionné.");
        }
    } else {
        alert("Aucun dossier sélectionné.");
    }
} else {
    alert("Veuillez vérifier la taille du plan de travail avant de continuer.");
}
