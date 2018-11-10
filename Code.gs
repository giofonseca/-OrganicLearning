/**This code creates the structure of an #OrganicLearningResource something like an "interactive Mind Map" in Google Slides.
To start, user has to install the OrganicLearningResource Add-on. Once it is installed there will be some tools under
the respective menu. The first two tool are to subdivide an specific Slide into a number of new ones which will be automatically
interlinked with the mein one through a "Back Button Icon" and objects randomly positioned on the main Slide as links to each new slide.
There are two ways to do this at the moment, the first is by choosing a Drawing from Google Drive to be used as "Back button Icon".
The second one is by writing the URL from an image on the web to be used as "Back button Icon"

This code by Giovanni Fonseca Fonseca (@giofonsecaf) is licensed under a 
Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License.
Based on a work at http://organic-learning.education/
*/


function onOpen(e) {
  SlidesApp.getUi()
    //.createAddonMenu()
    .createMenu('OrganicLearningResource')
    .addItem('Back button from Google Drawings', 'openPickerDrawing')
    .addItem('Back button from an Image in Drive or Uploading one', 'openPickerImage')
    .addItem('Back button from image URL', 'openSidebar')
    .addSeparator()
    .addItem('Contact','contact')
    //.addSubMenu(SlidesApp.getUi().createMenu('My sub-menu')
           //.addItem('Sub-themes', 'numbersubtopics')
    .addToUi();
}

function onInstall(e) {
  SlidesApp.getUi()
    //.createAddonMenu()
    .createMenu('OrganicLearningResource')
    .addItem('Back button from Google Drawings', 'openPickerDrawing')
    .addItem('Back button from an Image in Drive or Uploading one', 'openPickerImage')
    .addItem('Back button from image URL', 'openSidebar')
    .addSeparator()
    .addItem('Contact','contact')
    //.addSubMenu(SlidesApp.getUi().createMenu('My sub-menu')
           //.addItem('Sub-themes', 'numbersubtopics')
    .addToUi();
}

function openPickerDrawing() {
  var html = HtmlService.createTemplateFromFile('PickerDrawing').evaluate()
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SlidesApp.getUi().showModalDialog(html, 'Select a drawing for the Back Button');
}

function openPickerImage() {
  var html = HtmlService.createTemplateFromFile('PickerImage').evaluate()
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SlidesApp.getUi().showModalDialog(html, 'Select an image for the Back Button');
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function getDrawing(drawingId) {
  /*Logger.log(docInfo);
  Logger.log(drawingId);
  Logger.log(pictureUrl);
  Logger.log(objTitle);*/
  var image = Drive.Files.get(drawingId);
  //Logger.log(image);
  var imageBlob = getBlob(image.exportLinks['image/png']);
  
  var ui = SlidesApp.getUi(); // Same variations.
  
  var result = ui.prompt(
    'What Slide do you want to subdivide?',
    ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var mainSlide = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('The slide to be subdivided is: ' + mainSlide + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get the number of the slide to subdivide.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
  
  var result = ui.prompt(
    'How many new slides do you want me to create?',
    ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var numSlides = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('You want ' + numSlides + ' new slides.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get the amount of new slides that you want.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
  
  //to declare the variable where the names of the subtopics will be saved
  var subtopicsNames = [];
  
  // Asking user for the name of each subtopic
  for (i = 0; i < numSlides; i++) {
    var index= i+1;
    var response = ui.prompt('Name of the sub-topics', 'What is the name of subtopic number '+ index +' ?', ui.ButtonSet.OK_CANCEL);
    
    // getting the name of each subtopic into the array subtopicsNames
    if (response.getSelectedButton() == ui.Button.OK) {
      Logger.log('The subtopic number' + index + 'to add is %s.', response.getResponseText());
      subtopicsNames [i] = response.getResponseText();
      
      // confirm user the entered name of each subtopic
      ui.alert('OK, we added ' + subtopicsNames[i] + ' to the list!');
      
      //in case user want to cancel...or closing the window
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('The user didn\'t want to provide the information.');
      break;
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
      break;
    }  
  } 
  mainSlide = mainSlide -1;
  Logger.log(' '+subtopicsNames+' ',' '+mainSlide+' ');
  for (i = 0; i<subtopicsNames.length; i++) {
    
    // getting the current presentation
    var presentation= SlidesApp.getActivePresentation();  
    var selection = SlidesApp.getActivePresentation().getSelection();
    
    
    // defining the "Menu" slide as the first one [0] where the buttonGoTo will be
    var menuSlide = presentation.getSlides()[mainSlide];
    
    //creating a new slide with Title Only layout
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    
    //setting the title of the slide with the respective subtopic
    var placeholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    placeholder.asShape().getText().setText(subtopicsNames[i]);
    
    //move the inserted slide i+1 positions under the Mind Map Slide
    slide.move(mainSlide+i+1);
    
    // two random numbers for the position of the buttonGoTo
    var a = getRandomInt(0,650);
    var b = getRandomInt(0,350);
    
    // insert the buttonGoTo (round rectangle shape) on a random position at the "Menu" slide
    var buttonGoTo = menuSlide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, a, b, 100, 30);
    buttonGoTo.scaleHeight(1.5);
    buttonGoTo.getFill().setTransparent();
    // buttonGoTo.
    // setting the text of the buttonGoTo according to the respective subtopic
    buttonGoTo.getText().setText(subtopicsNames[i]);
    
    // linking the buttonGoTo to the respective Slide on the presentation (current slide) 
    buttonGoTo.setLinkSlide(slide);
    
    // inserting the URL image chosen by user to be use as a back-button to go to the Mind Map Slide
    var position = {left: 650, top: 345};
    var size = {width: 50, height: 50};
    var image2= slide.insertImage(imageBlob, position.left, position.top, size.width, size.height);
    
    // Linking the back-button to the first slide. 
    image2.setLinkSlide(menuSlide);
  }
}

function getImage(imageId) {
  /*Logger.log(docInfo);
  Logger.log(drawingId);
  Logger.log(pictureUrl);
  Logger.log(objTitle);*/
  var image = DriveApp.getFileById(imageId);
  var imageBlob = image.getBlob();
  var ui = SlidesApp.getUi(); // Same variations.
  
  var result = ui.prompt(
    'What Slide do you want to subdivide?',
    ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var mainSlide = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('The slide to be subdivided is: ' + mainSlide + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get the number of the slide to subdivide.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
  
  var result = ui.prompt(
    'How many new slides do you want me to create?',
    ui.ButtonSet.OK_CANCEL);
  
  // Process the user's response.
  var button = result.getSelectedButton();
  var numSlides = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('You want ' + numSlides + ' new slides.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get the amount of new slides that you want.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
  
  //to declare the variable where the names of the subtopics will be saved
  var subtopicsNames = [];
  
  // Asking user for the name of each subtopic
  for (i = 0; i < numSlides; i++) {
    var index= i+1;
    var response = ui.prompt('Name of the sub-topics', 'What is the name of subtopic number '+ index +' ?', ui.ButtonSet.OK_CANCEL);
    
    // getting the name of each subtopic into the array subtopicsNames
    if (response.getSelectedButton() == ui.Button.OK) {
      Logger.log('The subtopic number' + index + 'to add is %s.', response.getResponseText());
      subtopicsNames [i] = response.getResponseText();
      
      // confirm user the entered name of each subtopic
      ui.alert('OK, we added ' + subtopicsNames[i] + ' to the list!');
      
      //in case user want to cancel...or closing the window
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('The user didn\'t want to provide the information.');
      break;
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
      break;
    }  
  } 
  mainSlide = mainSlide -1;
  Logger.log(' '+subtopicsNames+' ',' '+mainSlide+' ');
  for (i = 0; i<subtopicsNames.length; i++) {
    
    // getting the current presentation
    var presentation= SlidesApp.getActivePresentation();  
    var selection = SlidesApp.getActivePresentation().getSelection();
    
    
    // defining the "Menu" slide as the first one [0] where the buttonGoTo will be
    var menuSlide = presentation.getSlides()[mainSlide];
    
    //creating a new slide with Title Only layout
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    
    //setting the title of the slide with the respective subtopic
    var placeholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    placeholder.asShape().getText().setText(subtopicsNames[i]);
    
    //move the inserted slide i+1 positions under the Mind Map Slide
    slide.move(mainSlide+i+1);
    
    // two random numbers for the position of the buttonGoTo
    var a = getRandomInt(0,650);
    var b = getRandomInt(0,350);
    
    // insert the buttonGoTo (round rectangle shape) on a random position at the "Menu" slide
    var buttonGoTo = menuSlide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, a, b, 100, 30);
    buttonGoTo.scaleHeight(1.5);
    buttonGoTo.getFill().setTransparent();
    // buttonGoTo.
    // setting the text of the buttonGoTo according to the respective subtopic
    buttonGoTo.getText().setText(subtopicsNames[i]);
    
    // linking the buttonGoTo to the respective Slide on the presentation (current slide) 
    buttonGoTo.setLinkSlide(slide);
    
    // inserting the URL image chosen by user to be use as a back-button to go to the Mind Map Slide
    var position = {left: 650, top: 345};
    var size = {width: 50, height: 50};
    var image2= slide.insertImage(imageBlob, position.left, position.top, size.width, size.height);
    
    // Linking the back-button to the first slide. 
    image2.setLinkSlide(menuSlide);
  }
}


/*function getBlob(url) {
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  return response.getBlob();
}*/

function getRandomInt(min, max) {
  return Math.floor(Math.random() * (max - min)) + min;
}

function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('New Sub-structure');
    
  SlidesApp.getUi().showSidebar(html);
}

// create the slides for each subtopic
function createSubSlides(menuIndex,backButtonIcon,subtopics) {
  
  // getting user interface 
  var ui = SlidesApp.getUi();  
  
  //to declare the variable where the names of the subtopics will be saved
  var subtopicsNames = [];
   
  // Asking user for the name of each subtopic
  for (i = 0; i < subtopics; i++) {
    var index= i+1;
    var response = ui.prompt('Name of the sub-topics', 'What is the name of subtopic number '+ index +' ?', ui.ButtonSet.OK_CANCEL);
    
    // getting the name of each subtopic into the array subtopicsNames
    if (response.getSelectedButton() == ui.Button.OK) {
      Logger.log('The subtopic number' + index + 'to add are %s.', response.getResponseText());
      subtopicsNames [i] = response.getResponseText();
      
      // confirm user the entered name of each subtopic
      ui.alert('OK, we added ' + subtopicsNames[i] + ' to the list!');
      
      //in case user want to cancel...or closing the window
    } else if (response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('The user didn\'t want to provide the information.');
      break;
    } else {
      Logger.log('The user clicked the close button in the dialog\'s title bar.');
      break;
    }  
  } 
  menuIndex = menuIndex -1;
  Logger.log(' '+subtopicsNames+' ',' '+menuIndex+' ');
  for (i = 0; i<subtopicsNames.length; i++) {
 
    // getting the current presentation
    var presentation= SlidesApp.getActivePresentation();  
    var selection = SlidesApp.getActivePresentation().getSelection();

    
    // defining the "Menu" slide as the first one [0] where the buttonGoTo will be
   var menuSlide = presentation.getSlides()[menuIndex];
    
    //creating a new slide with Title Only layout
    var slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_ONLY);
    
    //setting the title of the slide with the respective subtopic
    var placeholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    placeholder.asShape().getText().setText(subtopicsNames[i]);

    //move the inserted slide i+1 positions under the Mind Map Slide
    slide.move(menuIndex+i+1);
    
    // two random numbers for the position of the buttonGoTo
    var a = getRandomInt(0,650);
    var b = getRandomInt(0,350);
    
    // insert the buttonGoTo (round rectangle shape) on a random position at the "Menu" slide
    var buttonGoTo = menuSlide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, a, b, 100, 30);
    buttonGoTo.scaleHeight(1.5);
    buttonGoTo.getFill().setTransparent();
    
   // buttonGoTo.
    // setting the text of the buttonGoTo according to the respective subtopic
    buttonGoTo.getText().setText(subtopicsNames[i]);
    
    // linking the buttonGoTo to the respective Slide on the presentation (current slide) 
    buttonGoTo.setLinkSlide(slide);
    
    // inserting the URL image chosen by user to be use as a back-button to go to the Mind Map Slide
    var position = {left: 650, top: 345};
    var size = {width: 50, height: 50};
    var image= slide.insertImage(backButtonIcon, position.left, position.top, size.width, size.height);
    
    // Linking the back-button to the first slide. 
    image.setLinkSlide(menuSlide);
    
  } 
}

function contact(){
  var html = HtmlService.createHtmlOutputFromFile('contact')
    .setTitle('Contact');
  SlidesApp.getUi().showModalDialog(html, ' ')  
  //SlidesApp.getUi().showSidebar(html);
}

function getBlob(url) {
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  return response.getBlob();
}