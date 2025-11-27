function attachSheetLoader(fileInputSelector, sheetSelectSelector, sheetSectionSelector, enableSubmit) {
    $(fileInputSelector).change(function() {
        console.log("Fichier sélectionné sur", fileInputSelector);
        var fileInput = $(fileInputSelector)[0];
        var file = fileInput.files[0];
        if (!file) {
            console.log("Aucun fichier sélectionné pour", fileInputSelector);
            return;
        }

        var formData = new FormData();
        // IMPORTANT : le backend attend le champ "file"
        formData.append('file', file);

        $.ajax({
            url: '/get_sheets',
            type: 'POST',
            data: formData,
            processData: false,
            contentType: false,
            success: function(response) {
                console.log("Réponse du serveur pour", fileInputSelector, ":", response);
                if (response.error) {
                    alert("Erreur : " + response.error);
                    return;
                }

                var sheetSelect = $(sheetSelectSelector);
                sheetSelect.empty();

                response.sheet_names.forEach(function(sheet) {
                    sheetSelect.append($('<option>', {
                        value: sheet,
                        text: sheet
                    }));
                });

                $(sheetSectionSelector).show();

                if (enableSubmit) {
                    $('#submit_button').prop('disabled', false);
                }
            },
            error: function(xhr, status, error) {
                console.log("Erreur AJAX pour", fileInputSelector, ":", error);
                alert("Erreur lors de la lecture des feuilles : " + error);
            }
        });
    });
}

$(document).ready(function() {
    console.log("Document prêt.");

    // Fichier source : active aussi le bouton "Téléverser"
    attachSheetLoader('#file', '#sheet_name', '#sheet_selection', true);

    // Fichier cible : uniquement liste déroulante, sans toucher au bouton
    attachSheetLoader('#file_target', '#sheet_name_target', '#sheet_selection_target', false);
});