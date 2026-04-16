const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, LevelFormat,
  Header, Footer, ImageRun, VerticalAlign,
  HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType,
  PageNumber, PageNumberElement, PageNumberSeparator, PageBreak
} = require('docx');
const fs = require('fs');

const logoData = fs.readFileSync('/home/claude/logo.png');
const sseRechtsOben   = fs.readFileSync('/home/claude/sse_gelb_rechts_oben.png');
const sseLinksUnten   = fs.readFileSync('/home/claude/sse_gelb_links_unten.png');
const sseRechtsUnten  = fs.readFileSync('/home/claude/sse_gelb_rechts_unten.png');

// SSE floating image helper
// A4 page: 7,559,760 x 10,690,440 EMU | 8mm margin = 288,189 EMU
const SSE_PX = 20;
const SSE_EMU = SSE_PX * 9525;
const EDGE = 288189;
const PAGE_W = 7559760;
const PAGE_H = 10690440;

function sseFloat(data, hOffset, vOffset) {
  return new Paragraph({
    children: [new ImageRun({
      data,
      transformation: { width: SSE_PX, height: SSE_PX },
      type: "png",
      floating: {
        horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: hOffset },
        verticalPosition:   { relative: VerticalPositionRelativeFrom.PAGE,   offset: vOffset },
        allowOverlap: true,
        behindDocument: false,
        lockAnchor: true,
        wrap: { type: TextWrappingType.NONE, side: "bothSides" }
      }
    })]
  });
}

const DEEP_BLACK = "101921";
const LEMON_YELLOW = "e2e200";
const PLAIN_WHITE = "ffffff";
const STREET_GREY = "545860";
const LIGHT_BG = "f5f5f5";

const noBorder = { style: BorderStyle.NONE, size: 0, color: "ffffff" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function sectionHeading(text) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [
      new TableCell({
        shading: { fill: DEEP_BLACK, type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 180, right: 180 },
        borders: noBorders,
        children: [new Paragraph({
          keepNext: true,
          children: [
            new TextRun({ text, bold: true, color: PLAIN_WHITE, font: "Lucida Sans", size: 19 })
          ]
        })]
      })
    ]})]
  });
}

function body(text, spacingAfter = 60) {
  return new Paragraph({
    spacing: { after: spacingAfter },
    children: [new TextRun({ text, font: "Lucida Sans", size: 17, color: DEEP_BLACK })]
  });
}

function bullet(text, spacingAfter = 50) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: spacingAfter },
    children: [new TextRun({ text, font: "Lucida Sans", size: 17, color: DEEP_BLACK })]
  });
}

function spacer(after = 80) {
  return new Paragraph({ spacing: { after }, children: [] });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2013",
        alignment: AlignmentType.LEFT,
        style: {
          paragraph: { indent: { left: 640, hanging: 320 } },
          run: { color: LEMON_YELLOW, bold: true, font: "Lucida Sans" }
        }
      }]
    }]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1700, right: 1134, bottom: 1400, left: 1134 }
      }
    },

    headers: {
      default: new Header({
        children: [
          new Table({
            width: { size: 9026, type: WidthType.DXA },
            columnWidths: [3200, 5826],
            rows: [new TableRow({ children: [
              new TableCell({
                borders: { top: noBorder, bottom: noBorder, left: noBorder, right: { style: BorderStyle.NONE, size: 0, color: "ffffff" } },
                margins: { top: 60, bottom: 60, left: 0, right: 200 },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  children: [new ImageRun({
                    data: logoData,
                    transformation: { width: 180, height: 99 },
                    type: "png"
                  })]
                })]
              }),
              new TableCell({
                borders: noBorders,
                margins: { top: 60, bottom: 60, left: 200, right: 0 },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [
                    new TextRun({ text: "IT-Richtlinie f\u00fcr Mitarbeiter", font: "Lucida Sans", size: 22, bold: true, color: DEEP_BLACK }),
                    new TextRun({ break: 1, text: "In Verbindung mit der DSGVO", font: "Lucida Sans", size: 16, color: STREET_GREY }),
                    new TextRun({ break: 1, text: "Version 2, g\u00fcltig ab 01.04.2026", font: "Lucida Sans", size: 14, color: STREET_GREY })
                  ]
                })]
              })
            ]})]
          }),
          // Lemon Yellow rule
          new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: LEMON_YELLOW, space: 1 } },
            spacing: { after: 0, before: 60 },
            children: []
          }),
          // SSE corners (floating, page-relative, appear on every page)
          sseFloat(sseRechtsOben,  PAGE_W - EDGE - SSE_EMU, EDGE),           // oben rechts
          sseFloat(sseLinksUnten,  EDGE,                    PAGE_H - EDGE - SSE_EMU), // unten links
          sseFloat(sseRechtsUnten, PAGE_W - EDGE - SSE_EMU, PAGE_H - EDGE - SSE_EMU), // unten rechts
        ]
      })
    },

    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: STREET_GREY, space: 4 } },
            spacing: { before: 80, after: 0 },
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: "Spedination GmbH \u00b7 Sportplatzweg 5a, A-6336 Langkampfen  |  Spedination Zillertal GmbH \u00b7 Sportplatzweg 3, A-6270 Uderns",
              font: "Lucida Sans", size: 14, color: STREET_GREY
            }),
            new TextRun({ break: 1, text: "Spedination Deutschland GmbH \u00b7 Am Neugrund 39, D-83088 Kiefersfelden  |  Spedination Poland Sp. z o.o. \u00b7 ul. Kongresowa 28, PL-25-672 Kielce  |  www.spedination.com", font: "Lucida Sans", size: 14, color: STREET_GREY })]
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            spacing: { before: 40, after: 0 },
            children: [
              new TextRun({ text: "Seite ", font: "Lucida Sans", size: 14, color: STREET_GREY }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Lucida Sans", size: 14, color: STREET_GREY }),
              new TextRun({ text: " von ", font: "Lucida Sans", size: 14, color: STREET_GREY }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Lucida Sans", size: 14, color: STREET_GREY }),
            ]
          })
        ]
      })
    },

    children: [

      // Dienstgeber/Dienstnehmer Block
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [9026],
        rows: [new TableRow({ children: [
          new TableCell({
            shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
            borders: {
              top: noBorder, bottom: noBorder, right: noBorder,
              left: { style: BorderStyle.SINGLE, size: 24, color: LEMON_YELLOW }
            },
            margins: { top: 140, bottom: 140, left: 220, right: 200 },
            children: [
              new Paragraph({ spacing: { after: 60 }, children: [
                new TextRun({ text: "Dienstgeber:", font: "Lucida Sans", size: 18, bold: true, color: DEEP_BLACK }),
                new TextRun({ text: "  Spedination GmbH, Sportplatzweg 5a, A-6336 Langkampfen", font: "Lucida Sans", size: 18, color: DEEP_BLACK })
              ]}),
              new Paragraph({ spacing: { after: 60 }, children: [
                new TextRun({ text: "Dienstnehmer:", font: "Lucida Sans", size: 18, bold: true, color: DEEP_BLACK }),
                new TextRun({ text: "  Berkant Aydin, Friedhofstra\u00dfe 8b / Top 6, 6300 W\u00f6rgl", font: "Lucida Sans", size: 18, color: DEEP_BLACK })
              ]}),
              new Paragraph({ spacing: { after: 0 }, children: [
                new TextRun({ text: "Geburtsdatum:", font: "Lucida Sans", size: 18, bold: true, color: DEEP_BLACK }),
                new TextRun({ text: "  05.03.2002 in Kufstein", font: "Lucida Sans", size: 18, color: DEEP_BLACK })
              ]}),
            ]
          })
        ]})]
      }),

      spacer(100),

      sectionHeading("Einleitung"),
      spacer(60),
      body("IT-Sicherheit geht uns alle an!"),
      body("In fast jedem Unternehmen werden mittlerweile Daten vorwiegend elektronisch verarbeitet. Die verarbeiteten Daten reichen von Kundendaten, personenbezogenen Daten, \u00fcber Finanzdaten bis hin zu besonders sch\u00fctzenswerten Daten. Viele Unternehmen sind dar\u00fcber hinaus mit Daten konfrontiert, die keinesfalls in H\u00e4nde Dritter fallen d\u00fcrfen \u2013 sei es aus Gr\u00fcnden des Datenschutzes oder weil es sich um vertrauliche Unternehmensdaten handelt."),
      body("Datensicherheit im Allgemeinen und speziell IT-Sicherheit sind daher unverzichtbar f\u00fcr den Unternehmenserfolg. Unternehmensdaten m\u00fcssen bestm\u00f6glich gesch\u00fctzt werden \u2013 sowohl gegen Aussp\u00e4hung als auch gegen Datenverlust durch technische Gebrechen."),
      body("Die nachfolgenden Punkte sind sowohl f\u00fcr das Unternehmen, in dem Sie arbeiten, von gro\u00dfer Bedeutung, aber auch Sie pers\u00f6nlich profitieren privat von dieser Richtlinie."),
      body("Mit Ihrer Unterschrift best\u00e4tigen Sie, dass Sie diese Richtlinien beachten und umsetzen werden.", 240),

      sectionHeading("Sicherer Umgang mit personenbezogenen Daten"),
      spacer(60),
      body("Personenbezogene Daten sind all jene Informationen, die sich auf eine nat\u00fcrliche Person beziehen und so R\u00fcckschl\u00fcsse auf deren Pers\u00f6nlichkeit erlauben. Besondere personenbezogene Daten (ethnische Herkunft, Gesundheit, Religion, Sexualit\u00e4t etc.) sind besonders sch\u00fctzenswert."),
      body("Das Speichern und Verarbeiten von personenbezogenen Daten ist nur unter Zustimmung des Betroffenen zul\u00e4ssig."),
      body("Bitte beachten Sie folgende Punkte:"),
      bullet("Personenbezogene Daten m\u00fcssen geheim gehalten werden. Nur bei schriftlicher Zustimmung d\u00fcrfen diese an Dritte weitergegeben werden."),
      bullet("Bei Weitergabe muss auf einen sicheren Kommunikationsweg geachtet werden. Ein unverschl\u00fcsseltes E-Mail erf\u00fcllt diese Anforderung NICHT."),
      bullet("Nach dem Ausscheiden d\u00fcrfen Sie personenbezogene Daten, die Ihnen beruflich zug\u00e4nglich gemacht wurden, nicht weitergeben oder anderweitig nutzen.", 240),

      sectionHeading("Social Media"),
      spacer(60),
      body("Soziale Medien sind f\u00fcr Unternehmen in Punkto Sicherheit ein zunehmendes Problem. JEDE Information \u2013 sei sie noch so unwichtig \u2013 kann f\u00fcr Dritte verwertbar sein. Ein Foto vom Arbeitsplatz kann Kundennamen zeigen; die Information \u00fcber ein Firmenwochenende kann einem Angreifer ein Zeitfenster bieten."),
      body("Bitte ber\u00fccksichtigen Sie folgende Punkte:"),
      bullet("Posten Sie keine Fotos von Ihrem Arbeitsplatz."),
      bullet("Posten Sie keine Statusinformationen, die das Unternehmen betreffen."),
      bullet("Geben Sie in Foren oder sozialen Medien keine Informationen \u00fcber das Unternehmen preis."),
      bullet("Verwenden Sie Pseudonyme f\u00fcr notwendige Fragen in Foren, die das Unternehmen betreffen."),
      bullet("Nennen Sie keine Namen \u2013 weder Ihren eigenen noch den des Unternehmens.", 240),

      sectionHeading("Clear Desk Policy"),
      spacer(60),
      body("Alle vertraulichen Dokumente am Arbeitsplatz m\u00fcssen so verwahrt werden, dass Unberechtigte (Reinigungspersonal, Besucher, unbefugte Kollegen) keinen Zugriff darauf erhalten."),
      body("Bitte beachten Sie folgende Punkte:"),
      bullet("Bei Verlassen des Arbeitsplatzes m\u00fcssen alle Ausdrucke und Kopien in verschlie\u00dfbaren Beh\u00e4ltnissen verstaut werden."),
      bullet("Lassen Sie keine Ausdrucke im Drucker/Kopierer liegen."),
      bullet("Bewahren Sie unter keinen Umst\u00e4nden Passwortnotizen am Arbeitsplatz auf."),
      bullet("Sperren Sie Ihren Computer beim Verlassen des Arbeitsplatzes (Windows-Taste + L)!", 240),

      sectionHeading("Pers\u00f6nliche Passw\u00f6rter"),
      spacer(60),
      body("Passw\u00f6rter sch\u00fctzen vor unbefugtem Zutritt zu Ihren Systemen und Daten \u2013 \u00e4hnlich wie ein Schl\u00fcssel zu Ihrer Wohnung."),
      body("Wichtig: Sie werden niemals per E-Mail, Telefon oder Chat aufgefordert, ein Passwort einzugeben oder zu \u00e4ndern. Sollte das je notwendig sein, wird es Ihnen pers\u00f6nlich mitgeteilt."),
      body("Bitte beachten Sie folgende Punkte:"),
      bullet("Verwenden Sie nie dasselbe Passwort f\u00fcr unterschiedliche Zug\u00e4nge."),
      bullet("Mindestens 8 Zeichen, bestehend aus Gro\u00df- und Kleinbuchstaben, einer Ziffer und einem Sonderzeichen."),
      bullet("Keine Namen, Geburtsdaten oder Telefonnummern verwenden \u2013 diese werden bei Angriffen zuerst ausprobiert."),
      bullet("Keine W\u00f6rterbuchbegriffe verwenden (auch nicht in anderen Sprachen)."),
      bullet("Trivial-Passw\u00f6rter (1234, hallohallo, abcdefgh etc.) sind ungeeignet."),
      bullet("Geben Sie Ihr Passwort niemandem weiter \u2013 auch nicht an Kollegen oder IT-Betreuung."),
      bullet("Ihr initiales Passwort erhalten Sie ausschlie\u00dflich pers\u00f6nlich von der IT-Abteilung."),
      bullet("Tipp: Merken Sie sich einen Satz und verwenden Sie die Anfangsbuchstaben \u2013 z.B. \u201eDie Arbeit beginnt jeden Tag um 7 Uhr\u201c \u2192 DAbjTu7U"),
      bullet("Verdacht auf Kompromittierung? Sofort die IT-Abteilung kontaktieren.", 240),

      sectionHeading("Zugangsdaten von Web-Portalen"),
      spacer(60),
      body("Zugangsdaten f\u00fcr Web-Portale m\u00fcssen sicher gespeichert werden \u2013 E-Mail oder Dateiablage ist unzul\u00e4ssig. Empfohlen wird ein digitaler Passwortsafe. Bitte setzen Sie sich mit Ihrer IT-Abteilung in Verbindung."),
      body("Zugangsdaten d\u00fcrfen nach Austritt unter keinen Umst\u00e4nden gespeichert oder weitergenutzt werden.", 240),

      sectionHeading("Dokumente richtig entsorgen"),
      spacer(60),
      body("Sorglos weggeworfene Dokumente stellen ein ernstes Sicherheitsproblem dar. Alle vertraulichen Dokumente m\u00fcssen sicher entsorgt werden \u2013 mittels Dokumenten-Schredder oder einem spezialisierten Entsorgungsdienstleister."),
      bullet("Werfen Sie wichtige Dokumente niemals in den Papierkorb \u2013 auch nicht Archivmaterial.", 240),

      sectionHeading("Speicherung von Daten und Cloud-Dienste"),
      spacer(60),
      body("Daten d\u00fcrfen ausschlie\u00dflich auf den daf\u00fcr vorgesehenen, von der IT-Abteilung freigegebenen Systemen gespeichert werden (Netzlaufwerk oder Dokumentenmanagement). Lokale Laufwerke (C:) sind nicht zul\u00e4ssig."),
      body("Die Speicherung von Unternehmensdaten auf privaten Cloud-Diensten (z.B. Google Drive, Dropbox, private OneDrive-Konten o.\u00e4.) ist ebenfalls untersagt. Private Cloud-Dienste bieten keine Kontrolle dar\u00fcber, wo und wie Daten gespeichert oder verarbeitet werden.", 240),

      sectionHeading("Umgang mit mobilen IT-Ger\u00e4ten"),
      spacer(60),
      body("Mobile Ger\u00e4te (Notebooks, Smartphones) stellen erh\u00f6hte Sicherheitsrisiken dar und sind attraktive Ziele f\u00fcr Diebstahl."),
      bullet("Ger\u00e4t nicht unbeaufsichtigt lassen und keinen anderen Personen \u00fcberlassen."),
      bullet("Sichtschutz bei Passworteingabe beachten \u2013 wie am Bankomaten."),
      bullet("Privaten Cloud-Speicher nicht f\u00fcr Unternehmensdaten nutzen."),
      bullet("Nur von der IT-Abteilung freigegebene Apps installieren."),
      bullet("Diebstahl oder Verlust sofort der IT-Abteilung melden."),
      bullet("Datenvolumen im Blick behalten, um Mehrkosten zu vermeiden.", 240),

      sectionHeading("Internetnutzung"),
      spacer(60),
      body("Beim Surfen im Internet lauern Gefahren, die nicht immer sofort erkennbar sind. Es liegt in Ihrer Verantwortung, diese zu erkennen und entsprechend zu reagieren."),
      bullet("Gebrauchen Sie Ihren Hausverstand \u2013 betr\u00fcgerische E-Mails imitieren oft bekannte Anbieter."),
      bullet("Keine pers\u00f6nlichen Daten \u00fcber unsichere Verbindungen (kein HTTPS) \u00fcbermitteln."),
      bullet("Websites mit kostenloser Software oder Gewinnspielen grunds\u00e4tzlich misstrauen."),
      bullet("Download von Dateien kann lizenz- und urheberrechtliche Probleme verursachen."),
      bullet("Hackerseiten und Seiten mit gecr\u00e4ckter Software meiden."),
      bullet("Keine Websites mit pornografischen, gewaltverherrlichenden oder strafrechtlich bedenklichen Inhalten aufrufen."),
      bullet("Im Zweifel lieber einmal zu viel bei der IT-Abteilung nachfragen.", 240),

      sectionHeading("Protokollierung"),
      spacer(60),
      body("Jeder Datenverkehr unterliegt einer Protokollierung und Auswertung zur fr\u00fchzeitigen Erkennung von Datenverletzungen oder Schadcode. Die Auswertung erfolgt ausschlie\u00dflich in Verbindung mit der Gesch\u00e4ftsleitung und unter Wahrung des Datenschutzes.", 240),

      sectionHeading("SSL Interception"),
      spacer(60),
      body("Verschl\u00fcsselte Verbindungen erm\u00f6glichen es auch Schadsoftware, unentdeckt zu kommunizieren. Um dies zu verhindern, werden bestimmte Datenpakete zentral analysiert. Ausgenommen sind Finanzdienstleister, Beh\u00f6rden, Rechtsanw\u00e4lte, Gewerkschaften und medizinische Einrichtungen."),
      body("Mit Unterzeichnung dieser Richtlinie nimmt der Mitarbeiter zur Kenntnis, dass verschl\u00fcsselte Verbindungen im Rahmen der SSL Interception analysiert werden k\u00f6nnen.", 240),

      sectionHeading("E-Mail-Nutzung"),
      spacer(60),
      body("E-Mail ist ein bev\u00f6lkertes Angriffsziel. Spam-, Hoax- und Phishing-Mails machen ca. zwei Drittel des weltweiten E-Mail-Aufkommens aus."),
      bullet("\u00d6ffnen Sie keine E-Mails mit verd\u00e4chtigem Absender oder Betreff."),
      bullet("Keine verd\u00e4chtigen Dateinh\u00e4nge \u00f6ffnen \u2013 auch nicht bei scheinbar bekannten Absendern."),
      bullet("Phishing-Mails, die Bankdaten oder Passw\u00f6rter anfordern, sofort l\u00f6schen."),
      bullet("Links in E-Mails vor dem Klicken pr\u00fcfen \u2013 sichtbarer Text muss nicht mit dem Ziel \u00fcbereinstimmen."),
      bullet("Keine Spam-Mails beantworten \u2013 das best\u00e4tigt nur die G\u00fcltigkeit Ihrer Adresse."),
      bullet("Kollegen \u00fcber erkannte Phishing-Versuche informieren."),
      bullet("Bei Abwesenheit den Abwesenheitsassistenten aktivieren.", 240),

      sectionHeading("Social Engineering"),
      spacer(60),
      body("Social Engineering bezeichnet das gezielte Manipulieren von Personen zur Erlangung vertraulicher Informationen. Bekanntes Beispiel: der FACC-Angriff, bei dem per gef\u00e4lschter E-Mail mehrere Millionen Euro erbeutet wurden."),
      bullet("Bei ungew\u00f6hnlichen Anfragen per Telefon oder E-Mail skeptisch sein."),
      bullet("Au\u00dfergew\u00f6hnliche Auftr\u00e4ge nach M\u00f6glichkeit pers\u00f6nlich besprechen."),
      bullet("Keine vertraulichen Informationen per Telefon oder E-Mail weitergeben."),
      bullet("Bedenken Sie: Social Engineering bleibt oft lange unentdeckt.", 240),

      sectionHeading("Nutzung von KI-Tools"),
      spacer(60),
      body("Die Eingabe von Unternehmens-, Kunden- oder personenbezogenen Daten in externe KI-Dienste (z.B. ChatGPT, Google Gemini, Microsoft Copilot o.\u00e4.) ist untersagt. Diese Dienste k\u00f6nnen eingegebene Daten f\u00fcr das Training ihrer Modelle verwenden. Die Nutzung von durch die IT-Abteilung freigegebenen KI-Tools ist davon ausgenommen.", 240),

      sectionHeading("Private Nutzung der IT"),
      spacer(60),
      body("Die Nutzung der firmeneigenen IT f\u00fcr private Zwecke ist untersagt. Dies betrifft s\u00e4mtliche Ger\u00e4te und Kommunikationsmittel des Unternehmens:"),
      bullet("Computer und Laptops"),
      bullet("Firmenhandy"),
      bullet("Festnetztelefon"),
      bullet("Firmen-E-Mail"),
      bullet("Firmen-Messenger und sonstige Kommunikationsplattformen", 240),

      sectionHeading("Warnungen und Fehlermeldungen"),
      spacer(60),
      body("Warnungen oder Fehlermeldungen, die Sie nicht selbst verursacht haben oder nicht beheben k\u00f6nnen, sind unverz\u00fcglich der IT-Abteilung zu melden.", 240),

      sectionHeading("Wechselmedien"),
      spacer(60),
      body("Externe Datentr\u00e4ger (USB-Sticks, SD-Karten, externe Festplatten, CDs, DVDs, per USB angeschlossene Smartphones) k\u00f6nnen Schadsoftware enthalten und das gesamte Firmennetzwerk gef\u00e4hrden. Die Verwendung ist generell untersagt \u2013 Ausnahmen nur mit Genehmigung der IT-Abteilung.", 240),

      sectionHeading("Installation von Applikationen"),
      spacer(60),
      body("Die eigenst\u00e4ndige Installation von Applikationen ist untersagt \u2013 auf Windows-Ger\u00e4ten ebenso wie auf Smartphones und Tablets. Antr\u00e4ge sind schriftlich an die IT-Abteilung zu stellen.", 240),

      sectionHeading("Austritt aus dem Unternehmen"),
      spacer(60),
      body("Bei Austritt beh\u00e4lt sich der Arbeitgeber das Recht vor, E-Mail-Adressen des ausscheidenden Mitarbeiters weiter zu verwenden. S\u00e4mtliche Dokumente, IT-Equipment und Unterlagen sind unaufgefordert zu \u00fcbergeben. Der Arbeitgeber ist Inhaber des im Besch\u00e4ftigungsverh\u00e4ltnis erzeugten geistigen Eigentums."),
      body("Eine willk\u00fcrliche L\u00f6schung von Dokumenten, E-Mails oder sonstigen firmenrelevanten Daten ist untersagt.", 300),

      spacer(60),
      body("Ich habe das vollst\u00e4ndig durchgelesen und vollinhaltlich verstanden. Mit meiner Unterschrift best\u00e4tige ich mich daran zu halten.", 400),

      // Unterschriftstabelle
      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [3200, 400, 5426],
        rows: [new TableRow({ children: [
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 6, color: DEEP_BLACK }, left: noBorder, right: noBorder, bottom: noBorder },
            margins: { top: 80, bottom: 600, left: 0, right: 0 },
            children: [new Paragraph({ children: [new TextRun({ text: "Ort, Datum", font: "Lucida Sans", size: 18, color: STREET_GREY })] })]
          }),
          new TableCell({ borders: noBorders, children: [new Paragraph({ children: [] })] }),
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 6, color: DEEP_BLACK }, left: noBorder, right: noBorder, bottom: noBorder },
            margins: { top: 80, bottom: 600, left: 0, right: 0 },
            children: [new Paragraph({ children: [new TextRun({ text: "Name in Blockbuchstaben und Unterschrift", font: "Lucida Sans", size: 18, color: STREET_GREY })] })]
          }),
        ]})]
      }),

      spacer(300),

      // Änderungsprotokoll – immer neue Seite
      new Paragraph({
        pageBreakBefore: true,
        children: []
      }),

      // Änderungsprotokoll
      sectionHeading("\u00c4nderungsprotokoll"),
      spacer(60),

      new Table({
        width: { size: 9026, type: WidthType.DXA },
        columnWidths: [900, 1400, 5326, 1400],
        rows: [
          // Header row
          new TableRow({ children: [
            new TableCell({
              shading: { fill: "303030", type: ShadingType.CLEAR },
              borders: noBorders,
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "Version", font: "Lucida Sans", size: 17, bold: true, color: PLAIN_WHITE })] })]
            }),
            new TableCell({
              shading: { fill: "303030", type: ShadingType.CLEAR },
              borders: noBorders,
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "Datum", font: "Lucida Sans", size: 17, bold: true, color: PLAIN_WHITE })] })]
            }),
            new TableCell({
              shading: { fill: "303030", type: ShadingType.CLEAR },
              borders: noBorders,
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "\u00c4nderungen", font: "Lucida Sans", size: 17, bold: true, color: PLAIN_WHITE })] })]
            }),
            new TableCell({
              shading: { fill: "303030", type: ShadingType.CLEAR },
              borders: noBorders,
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "Verantwortlich", font: "Lucida Sans", size: 17, bold: true, color: PLAIN_WHITE })] })]
            }),
          ]}),
          // V1
          new TableRow({ children: [
            new TableCell({
              shading: { fill: "f5f5f5", type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "V1", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
            new TableCell({
              shading: { fill: "f5f5f5", type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "01.10.2020", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
            new TableCell({
              shading: { fill: "f5f5f5", type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "Erstversion", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
            new TableCell({
              shading: { fill: "f5f5f5", type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "T. Kogler", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
          ]}),
          // V2
          new TableRow({ children: [
            new TableCell({
              shading: { fill: PLAIN_WHITE, type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "V2", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
            new TableCell({
              shading: { fill: PLAIN_WHITE, type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "01.04.2026", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
            new TableCell({
              shading: { fill: PLAIN_WHITE, type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [
                new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Passwortregeln aktualisiert: Initialpasswort von IT, keine eigenst\u00e4ndige \u00c4nderung, nur pers\u00f6nliche Kommunikation", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
                new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Abschnitt \u201eVerschl\u00fcsselte Kommunikation\u201c gestrichen (veraltet)", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
                new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Abschnitt \u201eDokumente entsorgen\u201c: Datentr\u00e4ger entfernt", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
                new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Speicherung von Daten und Cloud-Dienste zusammengef\u00fchrt", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
                new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Private Nutzung erweitert: Handy, Festnetz, E-Mail, Messenger", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
                new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Neue Abschnitte: KI-Tools, SSL Interception (Einwilligung)", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
                new Paragraph({ spacing: { after: 0  }, children: [new TextRun({ text: "Sprachliche Korrekturen und CI-Formatierung", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] }),
              ]
            }),
            new TableCell({
              shading: { fill: PLAIN_WHITE, type: ShadingType.CLEAR },
              borders: { top: noBorder, bottom: { style: BorderStyle.SINGLE, size: 4, color: "dddddd" }, left: noBorder, right: noBorder },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [new TextRun({ text: "T. Kogler", font: "Lucida Sans", size: 17, color: DEEP_BLACK })] })]
            }),
          ]}),
        ]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/IT_Richtlinie_CI_final.docx', buffer);
  console.log('Done');
});
