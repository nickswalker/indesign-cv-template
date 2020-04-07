/*
  ExtendScript for Adobe Indesign to populate a text frame with information from a JSON CV

  Helpful resources:
  * https://github.com/ExtendScript/wiki/wiki
  * https://www.adobe.com/content/dam/acom/en/devnet/indesign/sdk/cs6/scripting/InDesign_ScriptingTutorial.pdf
  * https://wwwimages2.adobe.com/content/dam/acom/en/products/indesign/pdfs/InDesignCS5_ScriptingGuide_JS.pdf
  * http://yearbook.github.io/esdocs/#/InDesign/
  * https://www.indesignjs.de/extendscriptAPI/indesign-latest/#Application.html
*/

/* TODO: 
   * Author underline GREP style update based on profile data
   * Handle email and homepage link in upper right corner automatically
   * Add or remove pages to fit content automatically
   * Force author strings to use NBSP between initial for publications, presentations
*/
#target indesign
#targetengine "session"
#include "json2.js"
 //Turn user interaction on to allow for display of dialogs, alerts. etc.
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

myDocument = app.activeDocument;
paraStyles = {
    name: getParagraphStyle("Name"),
    sectionTitle: getParagraphStyle("Section Title"),
    sectionTitleFirst: getParagraphStyle("Section Title (First)"),
    dateBullet: getParagraphStyle("Date Bullet"),
    hangingBullet: getParagraphStyle("Hanging Bullet"),
    basic: myDocument.paragraphStyles.item("[Basic Paragraph]")
};
charStyles = {
    italic: getCharacterStyle("Italic"),
    icon: getCharacterStyle("Icon"),
    brandIcon: getCharacterStyle("Brand Icon"),
};

function getParagraphStyle(name) {
    try {
        style = myDocument.paragraphStyles.item(name);
    } catch (myError) {
        // The style did not exist, so create it.
        style = myDocument.characterStyles.add({
            "name": name
        });
    }
    return style;
}

function getCharacterStyle(name) {
    try {
        style = myDocument.characterStyles.item(name);
    } catch (myError) {
        // The style did not exist, so create it.
        style = myDocument.characterStyles.add({
            "name": name
        });
    }
    return style;
}

function appendParaBare(story, style) {
    story.insertionPoints[-1].contents += "\r";
    var paragraph = story.paragraphs[-1];
    if (style) {
        paragraph.applyParagraphStyle(style);
    } else {
        paragraph.applyParagraphStyle(paraStyles.basic)
    }
    return paragraph;
}

function appendPara(content, story, style) {
    var paragraph = appendParaBare(story, style);
    paragraph.contents = content + "\r";
    return paragraph;
}

function addText(content, insertionPoint, style) {
    var styleStartIndex = insertionPoint.index;
    var styleEndIndex = styleStartIndex + content.length - 1;
    var story = insertionPoint.parentStory;
    insertionPoint.contents += content;
    story.characters.itemByRange(styleStartIndex, styleEndIndex).appliedCharacterStyle = style;
    return styleEndIndex;
}

function addBrandIcon(content, insertionPoint) {
    var styleStartIndex = insertionPoint.index;
    var styleEndIndex = styleStartIndex + content.length - 1;
    var story = insertionPoint.parentStory;
    insertionPoint.contents += content;
    story.characters.itemByRange(styleStartIndex, styleEndIndex).appliedCharacterStyle = charStyles.brandIcon;
    return styleEndIndex;
}

function addIcon(content, insertionPoint) {
    var styleStartIndex = insertionPoint.index;
    var styleEndIndex = styleStartIndex + content.length - 1;
    var story = insertionPoint.parentStory;
    insertionPoint.contents += content;
    story.characters.itemByRange(styleStartIndex, styleEndIndex).appliedCharacterStyle = charStyles.icon;
    return styleEndIndex;
}

function addLink(content, link, story, style) {
    var hyperLink = addHyperlinkDestination(myDocument, link);

    var insertionPoint = story.insertionPoints[-1];
    var startIdx = insertionPoint.index;
    var endIdx = startIdx + content.length - 1;
    insertionPoint.contents += content;
    var characters = story.characters.itemByRange(startIdx, endIdx);
    var source = addHyperlinkTextSource(myDocument, characters);
    try {
        var link = app.activeDocument.hyperlinks.add(source, hyperLink);
    } catch (err) {
        alert(err)
    }
    if (style) {
        characters.appliedCharacterStyle = style;
    }
    return endIdx;
}

function setLinkRange(range, link) {
    var hyperLink = addHyperlinkDestination(myDocument, link);
    var source = addHyperlinkTextSource(myDocument, range);
    try {
        var link = app.activeDocument.hyperlinks.add(source, hyperLink);
    } catch (err) {
        alert(err)
    }
}

function appendDateBullet(specs, story) {
    var content = "";
    if (specs.startDate) {
        content += specs.startDate;
        if (specs.endDate) {
            content += "\u2013" + specs.endDate.substr(-2);
        } else {
            content += "\u2014"
        }
        content += "\t"
    } else if (specs.date) {
        content += specs.date + "\t";
    }
    content += specs.content;

    if (specs.extra) {
        content += " \u2013 " + specs.extra
    }
    var paragraph = appendPara(content, story, paraStyles.dateBullet);
    if (specs.link) {
        var hyperLink = addHyperlinkDestination(myDocument, specs.link);
        var characters = paragraph.characters.itemByRange(0, content.length - 1);
        var source = addHyperlinkTextSource(myDocument, characters);
        try {
            var link = app.activeDocument.hyperlinks.add(source, hyperLink);
        } catch (err) {
            alert(err)
        }
    }
    return paragraph;
}

function appendLabeledBullet(specs, story) {
    var content = "";
    if (specs.label) {
        content += specs.label + "\t";
    }
    content += specs.content;

    if (specs.extra) {
        content += " \u2013 " + specs.extra
    }
    var paragraph = appendPara(content, story, paraStyles.dateBullet);
    if (specs.link) {
        var hyperLink = addHyperlinkDestination(myDocument, specs.link);
        var characters = paragraph.characters.itemByRange(0, content.length - 1);
        var source = addHyperlinkTextSource(myDocument, characters);
        try {
            var link = app.activeDocument.hyperlinks.add(source, hyperLink);
        } catch (err) {
            alert(err)
        }
    }
    return paragraph;
}

function appendHangingBullet(content, story) {
    var content = "\u2022 " + content;
    return appendPara(content, story, paraStyles.hangingBullet);
}

function appendHangingBulletLinked(content, link, story) {
    var content = "\u2022 " + content;
    var para = appendPara(content, story, paraStyles.hangingBullet);
    setLinkRange(para.characters.itemByRange(0, -1), link)
    return para;
}

function main() {
    var resume;
    var resumeFile = File.openDialog("Select the CV JSON file.");

    var content;
    if (resumeFile != false) {
        resumeFile.open('r');
        content = resumeFile.read();
        resume = JSON.parse(content);
        resumeFile.close(); // always close files after reading
    } else {
        alert("Bah!");
        return;
    }

    // Set some file metadata
    with(myDocument.metadataPreferences) {
        author = resume["basics"]["name"];
        description = resume["basics"]["summary"];
    }

    // The template uses this variables in the headers
    setTextVariable("Name", resume["basics"]["name"])
    setTextVariable("Title", resume["basics"]["label"])
    var insertionPoint = app.selection[0];
    var mainStory = insertionPoint.parent;
    mainStory.contents = "";

    var namePara = appendPara("", mainStory, paraStyles.name)
    var nameVar = mainStory.textVariableInstances.add(LocationOptions.AFTER, namePara.insertionPoints[0]);
    nameVar.associatedTextVariable = myDocument.textVariables.item("Name");

    var titlePara = appendPara("", mainStory, null)
    var textVar = mainStory.textVariableInstances.add(LocationOptions.AFTER, titlePara.insertionPoints[0]);
    textVar.associatedTextVariable = myDocument.textVariables.item("Title");

    appendPara("Education", mainStory, paraStyles.sectionTitleFirst)
    var education = resume["education"]
    for (var i = 0; i < education.length; i++) {
        var item = education[i];
        item.content = item.institution
        item.extra = item.organization
        appendDateBullet(item, mainStory)
        appendHangingBullet(item["studyType"] + " " + item["area"], mainStory)
        var highlights = item["highlights"]
        if (!highlights) {
            continue;
        }
        for (var j = 0; j < highlights.length; j++) {
            var highlight = highlights[j]
            if (highlight instanceof String) {
                appendHangingBullet(highlight, mainStory)
            } else {
                appendHangingBulletLinked(highlight.description, highlight.link, mainStory)
            }
        }
    }


    appendPara("Conference", mainStory, paraStyles.sectionTitle)
    var publications = resume["publications"]
    var conference = [];
    var journal = [];
    var symposium_workshop = [];
    for (var i = 0; i < publications.length; i++) {
        var item = publications[i];
        if (item.type == "conference") {
            conference.push(item);
        } else if (item.type == "journal") {
            journal.push(item)
        } else if (item.type == "symposium" || item.type == "workshop") {
            symposium_workshop.push(item)
        }
    }

    for (var i = 0; i < conference.length; i++) {
        var item = conference[i];
        item.content = "\"" +item.name + ".\" " + item.authors + ". " + item.publisher + ". " + item.location + ", " + item.releaseDate
        item.label = "[c" + (conference.length - i) + "]"
        var bullet = appendLabeledBullet(item, mainStory)
        var startIdx = item.label.length + item.name.length + item.authors.length + 6
        var characters = bullet.characters.itemByRange(startIdx, startIdx + item.publisher.length)
        characters.appliedCharacterStyle = charStyles.italic

    }

    appendPara("Journal", mainStory, paraStyles.sectionTitle)
    for (var i = 0; i < journal.length; i++) {
        var item = journal[i];
        item.content = "\"" + item.name + ".\" " + item.authors + ". " + item.publisher + ". " + item.releaseDate
        item.label = "[j" + (journal.length - i) + "]"
        var bullet = appendLabeledBullet(item, mainStory)
        var startIdx = item.label.length + item.name.length + item.authors.length + 6
        var characters = bullet.characters.itemByRange(startIdx, startIdx + item.publisher.length)
        characters.appliedCharacterStyle = charStyles.italic
    }

    appendPara("Refereed Symposium, Workshop", mainStory, paraStyles.sectionTitle)
    for (var i = 0; i < symposium_workshop.length; i++) {
        var item = symposium_workshop[i];
        item.content = "\"" + item.name + ".\" " + item.authors + ". " + item.publisher + ". " + item.location + ", " + item.releaseDate
        item.label = "[w" + (symposium_workshop.length - i) + "]"
        var bullet = appendLabeledBullet(item, mainStory)
        var startIdx = item.label.length + item.name.length + item.authors.length + 6
        var characters = bullet.characters.itemByRange(startIdx, startIdx + item.publisher.length)
        characters.appliedCharacterStyle = charStyles.italic
    }

    appendPara("Presentations", mainStory, paraStyles.sectionTitle)
    var presentations = resume["presentations"]
    for (var i = 0; i < presentations.length; i++) {
        var item = presentations[i];
        item.content = item.name + ". " + item.authors + ". " + item.venue + ". " + item.location + ". " + item.format + "."
        if (item.note) {
            item.content += " " + item.note + "."
        }
        appendDateBullet(item, mainStory)
    }

    appendPara("Recognition", mainStory, paraStyles.sectionTitle)
    var recognition = resume["awards"]
    for (var i = 0; i < recognition.length; i++) {
        var item = recognition[i];
        item.content = item.title
        item.extra = item.awarder
        appendDateBullet(item, mainStory)
    }

    appendPara("Research Competitions", mainStory, paraStyles.sectionTitle)
    var competitions = resume["competitions"]
    for (var i = 0; i < competitions.length; i++) {
        var item = competitions[i];
        item.content = item.result
        item.extra = item.name
        appendDateBullet(item, mainStory)
    }

    appendPara("Research Affiliations", mainStory, paraStyles.sectionTitle)
    var affiliations = resume["affiliations"]
    for (var i = 0; i < affiliations.length; i++) {
        var item = affiliations[i];
        item.content = item.name
        item.extra = item.institution
        appendDateBullet(item, mainStory)
        if (item.summary) {
            appendHangingBullet(item["summary"], mainStory)
        }
    }

    appendPara("Outreach", mainStory, paraStyles.sectionTitle)
    var outreach = resume["outreach"]
    for (var i = 0; i < outreach.length; i++) {
        var item = outreach[i];
        item.content = item.position
        item.extra = item.organization
        item.link = item.website
        appendDateBullet(item, mainStory)
        appendHangingBullet(item["summary"], mainStory)
    }

    appendPara("Service", mainStory, paraStyles.sectionTitle)
    var serviceItems = resume["volunteer"]
    for (var i = 0; i < serviceItems.length; i++) {
        var item = serviceItems[i];
        item.content = item.position
        item.extra = item.organization
        item.link = item.website
        appendDateBullet(item, mainStory)
    }

    appendPara("Grants Received", mainStory, paraStyles.sectionTitle)
    var grants = resume["grants"]
    for (var i = 0; i < grants.length; i++) {
        var item = grants[i];
        item.content = item.name
        item.extra = item.awarder
        appendDateBullet(item, mainStory)
        if (item.summary) {
            appendHangingBullet(item["summary"], mainStory)
        }
    }

    appendPara("Meeting Participation", mainStory, paraStyles.sectionTitle)
    var meetings = resume["meetings"]
    for (var i = 0; i < meetings.length; i++) {
        var item = meetings[i];
        item.content = item.name + ", " + item.location
        appendDateBullet(item, mainStory)
    }

    appendPara("Work and Teaching Experience", mainStory, paraStyles.sectionTitle)
    var work = resume["work"]
    for (var i = 0; i < work.length; i++) {
        var item = work[i];
        item.content = item.position
        item.extra = item.company
        item.link = item.website
        appendDateBullet(item, mainStory)
        for (var j = 0; j < item.highlights.length; j++) {
            appendHangingBullet(item.highlights[j], mainStory)
        }
    }

    appendPara("Skills", mainStory, paraStyles.sectionTitle)
    var skills = resume["skills"]
    for (var i = 0; i < skills.length; i++) {
        var item = skills[i];
        item.content = "" + item.level + " with " + item.name.toLowerCase()
        item.extra = ""
        for (var j = 0; j < item.keywords.length; j++) {
            item.extra += item.keywords[j] + ", "
        }
        item.extra = item.extra.slice(-2)
        appendHangingBullet(item.content + " \u2013 " + item.extra, mainStory)
    }

    appendPara("Personal", mainStory, paraStyles.sectionTitle)
    var profiles = resume["basics"]["profiles"]
    var homePageURL = resume["basics"]["website"]
    var homePageNoProtocol = homePageURL.substr(8)
    var homePara = appendPara("\t" + homePageNoProtocol, mainStory, paraStyles.hangingBullet)
    addIcon("home", homePara.insertionPoints[0])
    setLinkRange(homePara.characters.itemByRange(0, -1), homePageURL)
    for (var i = 0; i < profiles.length; i++) {
        var profile = profiles[i]
        var profilePara = appendPara("\t" + profile.url.substr(8), mainStory, paraStyles.hangingBullet)
        var iconName = profile.network.toLowerCase()
        addBrandIcon(iconName, profilePara.insertionPoints[0])
        setLinkRange(profilePara.characters.itemByRange(0, -1), profile.url)
    }
}

function addHyperlinkTextSource(document, text) {
    var retVal = undefined;
    do {
        try {
            var parentId;
            var parentElement = text.characters.item(0).parent[0];
            var isCell = parentElement instanceof Cell;
            if (isCell) {
                parentId = parentElement.parent.id + "*" + parentElement.index;
            } else {
                parentId = text.characters.item(0).parentStory[0].id;
            }
            var fromIdx = text.characters.item(0).index[0];
            var toIdx = text.characters.item(text.characters.length - 1).index[0];
            var sourceCount = document.hyperlinkTextSources.length;
            var toRemove = [];
            for (var idx = 0; idx < sourceCount; idx++) {
                try {
                    var source = document.hyperlinkTextSources.item(idx);
                    var sourceText = source.sourceText;
                    var sourceElement = sourceText.characters.item(0).parent;
                    var sourceIsCell = sourceElement instanceof Cell;
                    var sourceId;
                    if (sourceIsCell) {
                        sourceId = sourceElement.parent.id + "*" + sourceElement.index;
                    } else {
                        sourceId = sourceText.parentStory.id;
                    }
                    if (sourceId == parentId) {
                        var sourceFromIdx = sourceText.characters.firstItem().index;
                        var sourceToIdx = sourceText.characters.lastItem().index;
                        if (sourceToIdx >= fromIdx && toIdx >= sourceFromIdx) {
                            toRemove.push(source);
                        }
                    }
                } catch (err) {}
            }

            for (var idx = 0; idx < toRemove.length; idx++) {
                toRemove[idx].remove();
            }

            retVal = document.hyperlinkTextSources.add(text);
        } catch (err) {}
    }
    while (false);

    return retVal;
}

function addHyperlinkDestination(document, url) {
    var retVal = undefined;
    do {
        try {
            var linkCount = document.hyperlinkURLDestinations.length;
            for (var idx = 0; idx < linkCount; idx++) {
                try {
                    var destination = document.hyperlinkURLDestinations.item(idx);
                    if (destination.destinationURL == url) {
                        retVal = destination;
                        break; // for
                    }
                } catch (err) {}
            }
            if (!retVal) {
                retVal = document.hyperlinkURLDestinations.add(url);
            }
        } catch (err) {}
    }
    while (false);

    return retVal;
}

function setTextVariable(name, value) {
    var atv = myDocument.textVariables;

    try {
        atv.itemByName(name).variableOptions.contents = value;
    } catch (err) {
        alert(err)
        var newTV = myDocument.textVariables.add({
            name: name,
            variableType: VariableTypes.CUSTOM_TEXT_TYPE
        });
        newTV.variableOptions.contents = value;
    }
    return true;
}

main();