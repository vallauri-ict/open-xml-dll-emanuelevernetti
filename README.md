# OpenXmlPlayground

OpenXmlPlayground è un progetto che permette capire il funzionamento dell'OpenXML nel nostro caso applicato a Microsoft Word e a Microsoft Excel, due software che fanno parte della suite di Microsoft Office.

All'interno della soluzione abbiamo due progetti: un Windows Forms e una libreria di classi.
In quest'ultima troviamo due classi: ClsWord e ClsExcel: la prima contiene dei metodi che consentono di andare a creare un documento di testo di Microsoft Word, mentre la seconda contiene dei metodi che vanno a impostare un foglio di calcolo di Microsoft Excel.

I due esempi sono facilmenti testabili utilizzando i due bottoni presenti sulla Form principale.

## ClsWord
### Questi sono i riferimenti necessari per il corretto funzionamento

```C#
using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
```
## Metodi
* **public static void AddStyle(MainDocumentPart mainPart, string styleId, string styleName, string fontName, int fontSize, string rgbColor, bool isBold, bool isItalic, bool isUnderlined)**: consente, ricevendo come parametro la MainPart del documento, di applicare una serie di stili.
* **public static Paragraph CreateParagraphWithStyle(string styleId, JustificationValues justification)**: crea un paragrafo, gli applica degli stili (nome dello stile e giustificazione) e poi lo restituisce.
* **public static void AddTextToParagraph(Paragraph paragraph, string content)**: riceve un paragrafo e una stringa di testo e va a scrivere nel paragrafo il contenuto della stringa.
* **public static void InsertPicture(WordprocessingDocument wordprocessingDocument, string fileName)**: riceve come parametri il documento su cui si sta lavorando e inserisce l'immagine che ricava dal parametro "fileName" contenente il path dell'immagine.
* **private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)**: questo metodo viene richiamato dal precedente e va effettivamente ad aggiungere l'immagine nel documento.

## ClsExcel
### Questi sono i riferimenti necessari per il corretto funzionamento

```C#
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
```
## Metodi
* **public static void CreateExcelFile<T>(List<T> data, string path)**: crea un foglio di calcolo di Microsoft Excel

Gli altri metodi presenti all'interno della classe servono per andare a creare le celle all'interno del foglio di calcolo e per andare ad inserire del testo di prova all'interno di esse. 


AUTORE: Emanuele Vernetti

Indirizzo mail per richieste e/o informazioni: e.vernetti.1033@vallauri.edu
