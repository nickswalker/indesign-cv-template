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
    h1: getParagraphStyle("Header 1"),
    h1First: getParagraphStyle("Header 1 (First)"),
    h2: getParagraphStyle("Header 2"),
    h3: getParagraphStyle("Header 3"),
    h3Subtitle: getParagraphStyle("Header 3 (With Subtitle)"),
    h3SubtitleDate: getParagraphStyle("Header 3 (With Subtitle, Date Lead)"),
    dateBullet: getParagraphStyle("Date Bullet"),
    publicationBullet: getParagraphStyle("Publication Bullet"),
    bullet: getParagraphStyle("Basic Bullet"),
	indentedBullet: myDocument.paragraphStyles.item("Indented Bullet"),
    bulletYears: getParagraphStyle("Bullet (Year List)"),
    basic: myDocument.paragraphStyles.item("[Basic Paragraph]"),
    indented: myDocument.paragraphStyles.item("Indented")
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

function makeDateString(startDate, endDate) {
	var content = "";
	if (startDate) {
		content += startDate;
		if (endDate) {
			content += "\u2013" + endDate.substr(-2);
		} else {
			content += "\u2014"
		}
	}
	return content;
}

function appendDateBullet(specs, story) {
    var content = "";
    if (specs.startDate) {
        content += makeDateString(specs.startDate, specs.endDate)
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

function appendH3DateSubtitle(specs, story) {
	var content = "";
	if (specs.startDate) {
		content += makeDateString(specs.startDate, specs.endDate)
		content += "\t"
	} else if (specs.date) {
		content += specs.date + "\t";
	}
	content += specs.content;

	if (specs.extra) {
		content += " \u2013 " + specs.extra
	}
	var paragraph = appendPara(content, story, paraStyles.h3SubtitleDate);
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

function appendH3Subtitle(specs, story) {
	var content = "";
	content += specs.content;

	if (specs.extra) {
		content += " \u2013 " + specs.extra + "\t"
	}

	if (specs.startDate) {
		content += makeDateString(specs.startDate, specs.endDate)
	} else if (specs.date) {
		content += specs.date;
	}

	var paragraph = appendPara(content, story, paraStyles.h3Subtitle);
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
    var paragraph = appendPara(content, story, paraStyles.publicationBullet);
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
    //var content = "\u2022 " + content;
    return appendPara(content, story, paraStyles.indentedBullet);
}

function appendHangingBulletLinked(content, link, story) {
    //var content = "\u2022 " + content;
    var para = appendPara(content, story, paraStyles.indentedBullet);
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
        alert("Couldn't open CV JSON file");
        return;
    }

    // Set some file metadata
    myDocument.metadataPreferences.author = resume["basics"]["name"];
    myDocument.metadataPreferences.description = resume["basics"]["summary"];
    

    // The template uses these variables in the headers
    setTextVariable("Name", resume["basics"]["name"])
    setTextVariable("Title", resume["basics"]["label"])

    // Clear out anything in the main frame
    var insertionPoint = app.selection[0];
    var mainStory = insertionPoint.parent;
    mainStory.contents = "";

    var namePara = appendPara("", mainStory, paraStyles.name)
    var nameVar = mainStory.textVariableInstances.add(LocationOptions.AFTER, namePara.insertionPoints[0]);
    nameVar.associatedTextVariable = myDocument.textVariables.item("Name");

    var titlePara = appendPara("", mainStory, null)
    var textVar = mainStory.textVariableInstances.add(LocationOptions.AFTER, titlePara.insertionPoints[0]);
    textVar.associatedTextVariable = myDocument.textVariables.item("Title");


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

	var volunteerItems = resume["volunteer"]
	var serviceItems = []
	var reviewing = {}

	var venueWebsites = {};
	for (var i = 0; i < volunteerItems.length; i++) {
		var item = volunteerItems[i];
		if (item.position == "Reviewer") {
			if (!(item.shortName in reviewing)){
				reviewing[item.shortName] = []
			} 
			reviewing[item.shortName].push(item.date)
			if (item.website) {
				venueWebsites[item.shortName] = item.website
			}
		}
		else {
			serviceItems.push(item)
		}
	}
	var reviewingOrderedKeys = [];
	for (var key in reviewing) {
		reviewingOrderedKeys.push(key)
	}
	reviewingOrderedKeys.sort(function(a,b){
		var mostRecentA = reviewing[a][0]
		var mostRecentB = reviewing[b][0]
		var diff = mostRecentB - mostRecentA
		if (diff == 0) {
			diff = (a > b)
		}
		return diff
	})

    //var order = ["Education", "Conference","Journal","Refereed Symposium, Workshop", "Presentations", "Recognition","Outreach","Service", "Reviewing","Grants Received", "Meeting Participation", "Work and Teaching Experience", "Skills", "Personal"]
    var order = ["Education", "Conference","Journal","Refereed Symposium, Workshop", "Recognition","Outreach","Service", "Reviewing", "Work and Teaching Experience", "Skills", "Personal"]
    for (var i = 0; i < order.length; i++) {
    	var sectionName = order[i];
		appendPara(sectionName, mainStory, paraStyles.h1);
		switch (sectionName) {
			case "Education":
			var education = resume["education"]
			for (var j = 0; j < education.length; j++) {
			    var item = education[j];
			    item.content = item.institution
			    item.extra = item.location
			    appendH3DateSubtitle(item, mainStory)
			    appendHangingBullet(item["studyType"] + " " + item["area"], mainStory)
			    var highlights = item["highlights"]
			    if (!highlights) {
			        continue;
			    }
			    for (var k = 0; k < highlights.length; k++) {
			        var highlight = highlights[k]
			        if (highlight instanceof String) {
			            appendHangingBullet(highlight, mainStory)
			        } else {
			            appendHangingBulletLinked(highlight.description, highlight.link, mainStory)
			        }
			    }
			}
			break;
			case "Conference":

			for (var j = 0; j < conference.length; j++) {
			    var item = conference[j];
			    item.content = "\"" +item.name + ".\" " + item.authors + ". " + item.publisher + ". " + item.location + ", " + item.releaseDate
			    item.label = "[c" + (conference.length - j) + "]"
			    var bullet = appendLabeledBullet(item, mainStory)
			    var startIdx = item.label.length + item.name.length + item.authors.length + 6
			    var characters = bullet.characters.itemByRange(startIdx, startIdx + item.publisher.length)
			    characters.appliedCharacterStyle = charStyles.italic

			}
			break;
			case "Journal":
			for (var j = 0; j < journal.length; j++) {
			    var item = journal[j];
			    item.content = "\"" + item.name + ".\" " + item.authors + ". " + item.publisher + ". " + item.releaseDate
			    item.label = "[j" + (journal.length - j) + "]"
			    var bullet = appendLabeledBullet(item, mainStory)
			    var startIdx = item.label.length + item.name.length + item.authors.length + 6
			    var characters = bullet.characters.itemByRange(startIdx, startIdx + item.publisher.length)
			    characters.appliedCharacterStyle = charStyles.italic
			}
			break;
			case "Refereed Symposium, Workshop":
			for (var j = 0; j < symposium_workshop.length; j++) {
			    var item = symposium_workshop[j];
			    item.content = "\"" + item.name + ".\" " + item.authors + ". " + item.publisher + ". " + item.location + ", " + item.releaseDate
			    item.label = "[w" + (symposium_workshop.length - j) + "]"
			    var bullet = appendLabeledBullet(item, mainStory)
			    var startIdx = item.label.length + item.name.length + item.authors.length + 6
			    var characters = bullet.characters.itemByRange(startIdx, startIdx + item.publisher.length)
			    characters.appliedCharacterStyle = charStyles.italic
			}
			break;
			case "Presentations":
			var presentations = resume["presentations"]
			for (var j = 0; j < presentations.length; j++) {
			    var item = presentations[j];
			    item.content = item.name + ". " + item.authors + ". " + item.venue + ". " + item.location + ". " + item.format + "."
			    if (item.note) {
			        item.content += " " + item.note + "."
			    }
			    appendDateBullet(item, mainStory)
			}
			break;
			case "Recognition":
			var recognition = resume["awards"]
			for (var j = 0; j < recognition.length; j++) {
			    var item = recognition[j];
			    item.content = item.title
			    item.extra = item.awarder
			    appendDateBullet(item, mainStory)
			}
			break;
			case "Research Competitions":
			var competitions = resume["competitions"]
			for (var j = 0; j < competitions.length; j++) {
			    var item = competitions[j];
			    item.content = item.result
			    item.extra = item.name
			    appendDateBullet(item, mainStory)
			}
			break;
			case "Research Affiliations":
			var affiliations = resume["affiliations"]
			for (var j = 0; j < affiliations.length; j++) {
			    var item = affiliations[j];
			    item.content = item.name
			    item.extra = item.institution
			    appendDateBullet(item, mainStory)
			    if (item.summary) {
			        appendHangingBullet(item["summary"], mainStory)
			    }
			}
			break;
			case "Outreach":
			var outreach = resume["outreach"]
			for (var j = 0; j < outreach.length; j++) {
			    var item = outreach[j];
			    item.content = item.position
			    item.extra = item.organization
			    item.link = item.website
			    appendH3DateSubtitle(item, mainStory)
			    appendHangingBullet(item["summary"], mainStory)
			}
			break;
			case "Service":
			for (var j = 0; j < serviceItems.length; j++) {
			    var item = serviceItems[j];
			    item.content = item.position
			    item.extra = item.organization
			    item.link = item.website
			    appendDateBullet(item, mainStory)
			}
			break;
			case "Reviewing":
				var entries = ""
				var spans = []
			for (var j = 0; j < reviewingOrderedKeys.length; j++) {
				var key = reviewingOrderedKeys[j]
				var yearString = ""
				var years = reviewing[key]
			    for (var k = 0; k < years.length; k++) {
			        yearString += "\u0027" + years[k].slice(2) + ", "
			    }
			    yearString = yearString.slice(0, -2)
				toAdd = key + "\t" + yearString + "\n";
				spans.push([entries.length, entries.length + toAdd.length - 1])
				entries += toAdd;
			}
				// Drop the last newline character
				var organizationPara = appendPara(entries.substr(0,entries.length - 1), mainStory, paraStyles.bulletYears)
				for (var j = 0; j < spans.length; j++) {
					var key = reviewingOrderedKeys[j]
					var span = spans[j]
					setLinkRange(organizationPara.characters.itemByRange(span[0], span[1]), venueWebsites[key])

				}
			break;
			case "Grants Received":
			var grants = resume["grants"]
			for (var j = 0; j < grants.length; j++) {
			    var item = grants[j];
			    item.content = item.name
			    item.extra = item.awarder
			    appendDateBullet(item, mainStory)
			    if (item.summary) {
			        appendHangingBullet(item["summary"], mainStory)
			    }
			}
			break;
			case "Meeting Participation":
			var meetings = resume["meetings"]
			for (var j = 0; j < meetings.length; j++) {
			    var item = meetings[j];
			    item.content = item.name + ", " + item.location
			    appendDateBullet(item, mainStory)
			}
			break;
			case "Work and Teaching Experience":
			var work = resume["work"]
			for (var j = 0; j < work.length; j++) {
			    var item = work[j];
			    item.content = item.position
			    item.extra = item.company
			    item.link = item.website
			    appendH3DateSubtitle(item, mainStory)
			    for (var k = 0; k < item.highlights.length; k++) {
			        appendHangingBullet(item.highlights[k], mainStory)
			    }
			}
			break;
			case "Skills":
			var skills = resume["skills"]
			for (var j = 0; j < skills.length; j++) {
			    var item = skills[j];
			    item.content = "" + item.level + " with " + item.name.toLowerCase()
			    item.extra = ""
			    for (var k = 0; k < item.keywords.length; k++) {
			        item.extra += item.keywords[k] + ", "
			    }
			    item.extra = item.extra.slice(0, -2)
			    appendHangingBullet(item.content + " \u2013 " + item.extra, mainStory)
			}
			break;
			case "Personal":
			var profiles = resume["basics"]["profiles"]
			var homePageURL = resume["basics"]["website"]
			var homePageNoProtocol = homePageURL.substr(8)
			var homePara = appendPara("\u2003" + homePageNoProtocol, mainStory, paraStyles.indented)
			addIcon("home", homePara.insertionPoints[0])
			setLinkRange(homePara.characters.itemByRange(0, -1), homePageURL)
			for (var j = 0; j < profiles.length; j++) {
			    var profile = profiles[j]
				// Em space then URL of profile
			    var profilePara = appendPara("\u2003" + profile.url.substr(8), mainStory, paraStyles.indented)
			    var iconName = profile.network.toLowerCase()
			    addBrandIcon(iconName, profilePara.insertionPoints[0])
			    setLinkRange(profilePara.characters.itemByRange(0, -1), profile.url)
			}
			break;	

		}
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