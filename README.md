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
* AddStyle: consente, ricevendo come parametro la MainPart del documento, va ad applicare una serie di stili


AUTORE: Emanuele Vernetti

Indirizzo mail per richieste e/o informazioni: e.vernetti.1033@vallauri.edu
