function onOpen(e) {
  const ui = SlidesApp.getUi();
  ui
  .createAddonMenu()
  .addItem('Show sidebar', 'showSidebar')
  .addToUi();
  showSidebar();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const ui = SlidesApp.getUi();
  const sidebar = HtmlService.createHtmlOutput('<h1>Sizing</h1>' +
                                               '<button onClick="google.script.run.sameWidth();">Same width</button>' +
                                               '<button onClick="google.script.run.sameHeight();">Same height</button>' +
                                               '<button onClick="google.script.run.sameSize();">Same width and height</button>' +
                                               '<button onClick="google.script.run.increaseSize();">Increase size</button>' +
                                               '<button onClick="google.script.run.decreaseSize();">Decrease size</button>' +
                                               '<h1>Positioning</h1>' +
                                               '<button onClick="google.script.run.swapPositions();">Swap elements</button>' +
                                               '<button onClick="google.script.run.organizeAsGrid();">Organize as grid</button>' +
                                               '<h1>Alignment</h1>' +
                                               '<button onClick="google.script.run.alignLeft();">Align left</button>' +
                                               '<button onClick="google.script.run.alignCenter();">Align center</button>' +
                                               '<button onClick="google.script.run.alignRight();">Align right</button>' +
                                               '<button onClick="google.script.run.alignTop();">Align top</button>' +
                                               '<button onClick="google.script.run.alignMiddle();">Align middle</button>' +
                                               '<button onClick="google.script.run.alignBottom();">Align bottom</button>' +
                                               '<h1>Rotation</h1>' +
                                               '<button onClick="google.script.run.rotateClockwise();">Rotate clockwise</button>' +
                                               '<button onClick="google.script.run.rotateCounterclockwise();">Rotate counterclockwise</button>' +
                                               '<button onClick="google.script.run.flipVertically();">Flip vertically</button>' +
                                               '<button onClick="google.script.run.flipHorizontally();">Flip horizontally</button>' +
                                               '<h1>Text formatting</h1>' +
                                               '<button onClick="google.script.run.sameTextFormatting();">Same text formatting</button>' +
                                               '<button onClick="google.script.run.minimizeListIdention();">Minimize list indention</button>').setTitle("Sidebar title");
  ui.showSidebar(sidebar);
  console.log("Successfully added sidebar");
}

function getSelection() {
  return SlidesApp.getActivePresentation().getSelection();
}

function isPageElementSelection(selection) {
  return selection.getSelectionType() === SlidesApp.SelectionType.PAGE_ELEMENT
}

function getPageElements(selection) {
  return selection.getPageElementRange().getPageElements();
}

function roundNumber(num) {
  return Math.round(num * 100) / 100;
}

function sameWidth() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Set widths
  const targetWidth = pageElements[0].getWidth();
  for (let i=0; i<numElements; i++) {
    pageElements[i].setWidth(targetWidth);
  }
  
  // Debug
  console.log("Successfully applied width of %d to %d elements", roundNumber(targetWidth), numElements)
}

function sameHeight() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Set heights
  const targetHeight = pageElements[0].getHeight();
  for (let i=0; i<numElements; i++) {
    pageElements[i].setHeight(targetHeight);
  }
  
  // Debug  
  console.log("Successfully applied height of %d to %d elements", roundNumber(targetHeight), numElements)
}

function sameSize() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Set widths and heights  
  const targetHeight = pageElements[0].getHeight();
  const targetWidth = pageElements[0].getWidth();
  for (let i=0; i<numElements; i++) {
    pageElements[i].setHeight(targetHeight).setWidth(targetWidth);
  }
  
  console.log("Successfully applied height of %d and width of %d to %d elements", roundNumber(targetHeight), roundNumber(targetWidth), numElements)
}

function swapPositions() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements !== 2) {
    ui.alert("Please select exactly 2 elements");
    return;
  }
  
  // Save positions
  const [first, second] = pageElements;
  const firstLeft = first.getLeft();
  const firstTop = first.getTop();
  const secondLeft = second.getLeft();
  const secondTop = second.getTop();
  
  // Set positions  
  first.setLeft(secondLeft).setTop(secondTop);
  second.setLeft(firstLeft).setTop(firstTop);
  
  // Debug
  console.log("Successfully swapped 2 elements");
}

function getRight(element) {
  return element.getLeft() + element.getWidth();
}

function setRight(element, right) {
  return element.setLeft(right - element.getWidth());
}

function getBottom(element) {
  return element.getTop() + element.getHeight();
}

function setBottom(element, bottom) {
  return element.setTop(bottom - element.getHeight());
}

function alignLeft() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Find minimum left
  let targetLeft = pageElements[0].getLeft();
  let left = null;
  for (let i=1; i<numElements; i++) {
    left = pageElements[i].getLeft();
    if (left < targetLeft) {
      targetLeft = left;
    }
  }
  
  // Set lefts
  for (let i=0; i<numElements; i++) {
    pageElements[i].setLeft(targetLeft);
  }
  
  // Debug
  console.log("Successfully applied left position of %d to %d elements", roundNumber(targetLeft), numElements);
}

function alignRight() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Find maximum right
  let targetRight = getRight(pageElements[0]);
  let right = null;
  for (let i=1; i<numElements; i++) {
    right = getRight(pageElements[i])
    if (right > targetRight) {
      targetRight = right;
    }
  }
  
  // Set rights
  for (let i=0; i<numElements; i++) {
    setRight(pageElements[i], targetRight);
  }
  
  console.log("Successfully applied right position of %d to %d elements", roundNumber(targetRight), numElements);
}

function alignTop() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Find minimum top
  let targetTop = pageElements[0].getTop();
  let top = null;
  for (let i=1; i<numElements; i++) {
    top = pageElements[i].getTop();
    if (top < targetTop) {
      targetTop = top;
    }
  }
  
  // Set top
  for (let i=0; i<numElements; i++) {
    pageElements[i].setTop(targetTop);
  }
  
  // Debug
  console.log("Successfully applied top position of %d to %d elements", roundNumber(targetTop), numElements);
}

function alignBottom() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Find minimum bottom
  let targetBottom = getBottom(pageElements[0]);
  let bottom = null;
  for (let i=1; i<numElements; i++) {
    bottom = getBottom(pageElements[i])
    if (bottom > targetBottom) {
      targetBottom = bottom;
    }
  }
  
  // Set rights
  for (let i=0; i<numElements; i++) {
    setBottom(pageElements[i], targetBottom);
  }
  
  console.log("Successfully applied bottom position of %d to %d elements", roundNumber(targetBottom), numElements);
}

function alignCenter() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Find average center
  const avg = pageElements.reduce((accumulator, pageElement) => accumulator + pageElement.getLeft() + pageElement.getWidth() / 2, 0) / numElements;
  
  // Set left
  let element = null;
  for (let i=0; i<numElements; i++) {
    element = pageElements[i];
    element.setLeft(avg - element.getWidth() / 2);
  }
  
  // Debug
  console.log("Successfully applied center position to %d elements", numElements);
}

function alignMiddle() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Find average middle
  const avg = pageElements.reduce((accumulator, pageElement) => accumulator + pageElement.getTop() + pageElement.getHeight() / 2, 0) / numElements;
  
  // Set top
  let element = null;
  for (let i=0; i<numElements; i++) {
    element = pageElements[i];
    element.setTop(avg - element.getHeight() / 2);
  }
  
  // Debug
  console.log("Successfully applied middle position to %d elements", numElements);
}

function scaleElements(scaleWidth, scaleHeight) {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  for (let i=0; i<numElements; i++) {
    pageElements[i].scaleHeight(scaleHeight).scaleWidth(scaleWidth);
  }
  
  console.log("Successfully scaled size of %d elements by %d", numElements, scale);
}

function increaseSize() {
  scaleElements(1.1, 1.1)
}

function decreaseSize() {
  scaleElements(0.9, 0.9)
}

function flipVertically() {
  scaleElements(1, -1);
}

function flipHorizontally() {
  scaleElements(-1, 1);
}

function sameTextFormatting() {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  if (numElements <= 1) {
    ui.alert("Please select minimum 2 elements");
    return;
  }
  
  // Save text style
  const targetTextStyle = pageElements[0].asShape().getText().getTextStyle();
  const fontWeight = targetTextStyle.getFontWeight();
  const fontSize = targetTextStyle.getFontSize();
  const fontFamily = targetTextStyle.getFontFamily();
  const foregroundColor = targetTextStyle.getForegroundColor();
  const isItalic = targetTextStyle.isItalic();
  const isUnderline = targetTextStyle.isUnderline();
  const backgroundColor = targetTextStyle.getBackgroundColor();
  
  // Set text style
  for (let i=0; i<numElements; i++) {
    pageElements[i]
    .asShape()
    .getText()
    .getTextStyle()
    .setFontSize(fontSize)
    .setFontFamilyAndWeight(fontFamily, fontWeight)
    .setForegroundColor(foregroundColor)
    .setItalic(isItalic)
    .setUnderline(isUnderline)
    .setBackgroundColor(backgroundColor);
  }
  
  // Debug
  console.log("Successfully applied text formatting to %d elements", numElements)
}

function minimizeListIndention() {
  const ui = SlidesApp.getUi();
  ui.alert("Coming soon...");
}

function organizeInGrid() {
  const ui = SlidesApp.getUi();
  ui.alert("Coming soon...");
}

function rotate(rotation) {
  const ui = SlidesApp.getUi();
  const selection = getSelection();
  
  if (!isPageElementSelection(selection)) {
    ui.alert("Please select elements.");
    return;
  }
  
  const pageElements = getPageElements(selection);
  const numElements = pageElements.length;
  
  // Rotate
  let element = null;
  for (let i=0; i<numElements; i++) {
    element = pageElements[i];
    element.setRotation(element.getRotation() + rotation);
  }
  
  console.log("Successfully rotation of %d to %d elements", rotation, numElements)
}

function rotateClockwise() {
  rotate(45);
}

function rotateCounterclockwise() {
  rotate(-45);
}

function generateSlidesScript() {
  const ui = SlidesApp.getUi();
  const presentation = SlidesApp.getActivePresentation();
  const docTitle = presentation.getName() + ' Script';
  const slides = presentation.getSlides();
  const numSlides = slides.length;
  
  // Create document in the home Drive directory
  const speakerNotesDoc = DocumentApp.create(docTitle);
  console.log('Created document with id %s', speakerNotesDoc.getId());
  const docBody = speakerNotesDoc.getBody();
  const header = docBody.appendParagraph(docTitle);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  // Iterate through each slide and extract the speaker notes into the document body
  for (let i=0; i<numSlides; i++) {
    const section = docBody.appendParagraph(`Slide ${i + 1}`)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    const notes = slides[i].getNotesPage().getSpeakerNotesShape().getText().asString();
    docBody.appendParagraph(notes)
    .appendHorizontalRule();
  }
  
  ui.alert(`${speakerNotesDoc.getName()} has been created in your Drive files.`);
}
