# ConvertWordDoc2HtmlMultipleChoice

Convert a teachers Word document, for a multiple choice test , to a HTML page with radio buttons. This is accomplished by using a macro.
The Word document has to be structured like

Q1 Question1     

#1A answer A           
#1B answer B
#1C answer C  

Q2 Question2     

#2A answer A        {This is commment not sent to the html file}
#2B answer B
etc.


Questions and answers can be used with subscript and superscript , italic, underline etc., e.g. chemical formulas.
Text in curly brackets will not sent to the html file. So this can be used to add explanations.

In a previous version "When click on submit, an e-mail with an string of the given answers is ready to sent to the teacher."
In the last version when you click on the button the answers are copied to the clipboard. So the student can paste it in a Microsoft Form, or any other platform to share with the teacher. 

**You can find the Worddocument with the macro** in the WordDoc folder in this repository. The name of the file is Convert2MultiChoiceHtml.docm.

In an other repository I planned to publish an Outlook Add-In to evaluate the answers , to an Excel spreadsheet.

Credits goes to the creator of the word to html converter Toxaris

ConvertWordDoc2HtmlMultipleChoice is released under MIT open-source license. See the file "LICENSE.txt" for full licensing info.


