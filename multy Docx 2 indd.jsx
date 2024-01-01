#target "InDesign"

// Function to show folder picker dialog
function pickFolder() {
    var folder = Folder.selectDialog("Select a folder containing the Docx files");
    return folder;
}

/// Function to open an existing InDesign document, place the file, and save it
function processFile(file, templateFile) {
    // Open the existing InDesign document
    var doc = app.open(templateFile);

    // Loop through each page in the document
    for (var pageIndex = 0; pageIndex < doc.pages.length; pageIndex++) {
        // Create a text frame within the document margins
        var marginBounds = doc.pages[pageIndex].bounds;
        var textFrame = doc.pages[pageIndex].textFrames.add({
            geometricBounds: [
                marginBounds[0] + doc.marginPreferences.top,
                marginBounds[1] + doc.marginPreferences.left,
                marginBounds[2] - doc.marginPreferences.bottom,
                marginBounds[3] - doc.marginPreferences.right
            ]
        });

        // Place the file into the text frame
        textFrame.place(file);
    }

    // Save the document with the same name as the file
    var outputFolder = new Folder(file.parent + "/output");
    if (!outputFolder.exists) {
        outputFolder.create();
    }

    var outputFilePath = outputFolder + "/" + file.name.replace(/\.[^\.]+$/, ".indd");
    doc.save(new File(outputFilePath));

    // Close the document without saving changes (since it's already saved)
    doc.close(SaveOptions.no);
}


// Pick a folder using the dialog
var selectedFolder = pickFolder();

// Check if a folder was selected
if (selectedFolder) {
    var templateFile = File.openDialog("Select a blank InDesign file to serve as a template", "*.indd");

    // Check if a template file was selected
    if (templateFile) {
        var files = selectedFolder.getFiles();

        // Loop through each file in the selected folder
        for (var i = 0; i < files.length; i++) {
            if (files[i] instanceof File && files[i].name.match(/\.docx$/i)) {
                // Process each file
                processFile(files[i], templateFile);
            }
        }
    } else {
        $.writeln("No template file selected.");
    }
} else {
    $.writeln("No folder selected.");
}
