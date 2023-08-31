//Yearbook profiles txt reader and printer
//Hashim Iqbal
//20/05/2021

//Global vars
var file = new File;
var check = 0;

//UI

#targetengine "session";
var mainWindow = new Window("palette", "File Reader", undefined);
mainWindow.orientation = "column";

var groupOne = mainWindow.add("group", undefined, "groupOne");
groupOne.orientation = "row";
var fileLocBox = groupOne.add("edittext", undefined, "Selected File Location");
fileLocBox.size = [160, 20];
var getFileButton = groupOne.add("button", undefined, "File...");
getFileButton.helpTip = "Select a .txt file to fill the book with.";

var groupTwo = mainWindow.add("group", undefined, "groupTwo");
groupTwo.orientation = "row";
var applyButton = groupTwo.add("button", undefined, "Apply");

mainWindow.center();
mainWindow.show()

getFileButton.onClick = function() {
    file = file.openDlg("Please open a .txt file", "Acceptable Files:*.txt");
    fileLocBox.text = file.fsName;
    check = 1;
}

applyButton.onClick = function() {
    if(check == 0) {
        alert("Please select a file.");
        return false
    }

    else {
        var fileData;
        fileData = readTxt();
    }

    alert("Done");
}

function readTxt() {
    var txtObjects = [];
    file.open("r");

    var lines = file.read().toString();
    lines = lines.split('\n');
    //alert(lines);

    while(lines.length > 0) {

        txtObjects.push({
            Name: getField(lines),
            Nickname: getField(lines),
            YearJoined: getField(lines),
            MemorableMoments: getField(lines),
            BestQuotes: getField(lines),
            Friends: getField(lines),
            EmbarrassingMoments: getField(lines),
            InTenYears: getField(lines),
            Paragraph: getField(lines),
            Authors: getField(lines)
        })

        if(lines[0] == "") {
            lines.shift();
        }
    }

    file.close();
    printData(txtObjects);
}

function getField(lines) {
    return lines.shift();
}

function printData(data) {
    var openDocument = app.activeDocument;
    //var newDocument = app.documents.add();
    
    var pages = openDocument.pages;
    var layers = openDocument.layers; 
    
    var pageItems = openDocument.pageItems;
    
    alert(data.length);
    
    for(var x = 0; x < data.length; x++) {
        
            //var newPage = pages.add();
            
            var profileHeadingsTextFrame = pages[0].textFrames.add();
            profileHeadingsTextFrame.contents = "Also Known as: " + data[x].Nickname +  "\n" + "In custody since: " + data[x].YearJoined + "\n" + "Most wanted for: " + data[x].MemorableMoments + "\n" + "Most incriminating comments: " + data[x].BestQuotes + "\n" + "Partners in crime: " + data[x].Friends + "\n" + "Caught red handed: " + data[x].EmbarrassingMoments + "\n" + "Most likely to be doing in 10 years once released on parole: " + data[x].InTenYears + "\n" + "Statement for the defence: " + data[x].Paragraph + "\n" + "Profile authors: " + data[x].Authors;
            profileHeadingsTextFrame.geometricBounds = [28, 90, 114, 46];     
            
            var profilePrisonerNames = pages[0].textFrames.add();
            profilePrisonerNames.contents = data[x].Name;
            profilePrisonerNames.geometricBounds = [28, 50, 36, 92];
            
            var para1 = profilePrisonerNames.paragraphs[0];
            var para2 = profileHeadingsTextFrame.paragraphs[0];
            
            //Change text properties like font, styles and size.
            para1.appliedFont = app.fonts.item("Century"); 
            para1.pointSize = 10;

            para2. appliedFont = app.fonts.item("Century");
            para2.pointSize = 6;
            //para2.fontStyle = "Bold";

    }
}

/*
//var openDocument = app.activeDocument;

//var pages = openDocument.pages;
//var layers = openDocument.layers;

//var pageItems = openDocument.pageItems;

var newDocument = app.documents.add(2550, 3300, 300, "My new PS doc", NewDocumentMode.RGB);

var layers = newDocument.artLayers;
var newLayer = layers.add();

newLayer.name = "My new text layer";
newLayer.kind = LayerKind.TEXT;

var textItem = newLayer.textItem;

textItem.contents = readFile();

function readFile() {
    var inputFile = File("~/Desktop/Collyhill.txt");
    inputFile.encoding = "UTF-8";

    var inputData;
    var secondLine;

    inputFile.open("r");

    inputData = inputFile.read().toString();

    for(var i = 1; i <= 2; i++) {
        secondLine = inputFile.readln().toString();
    }

    inputFile.close();

    alert("Full Path: " + inputFile.fsName);
    alert("File Content: " + inputData);

    return secondLine;
}*/